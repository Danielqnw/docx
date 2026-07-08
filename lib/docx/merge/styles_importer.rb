# frozen_string_literal: true

require 'nokogiri'
require 'set'

module Docx
  module Merge
    class StylesImporter
      MAX_IMPORT_DEPTH = 10
      DEPENDENCY_ELEMENTS = %w[basedOn next link].freeze
      XML_NAMESPACES = Docx::Document::XML_NAMESPACES

      attr_reader :style_id_map

      def initialize(target_doc, source_doc)
        @target_doc = target_doc
        @source_doc = source_doc
        @style_id_map = {}
        @rename_sequence = 0
        @target_index = build_style_index(@target_doc.styles)
        @source_index = build_style_index(@source_doc.styles)
      end

      def import(source_style_id)
        import_style(source_style_id, depth: 0, visiting: Set.new)
      end

      private

      def build_style_index(styles_doc)
        return {} unless styles_doc

        styles_doc.xpath('//w:styles/w:style', XML_NAMESPACES).each_with_object({}) do |node, index|
          style_id = node['w:styleId']
          index[style_id] = node if style_id
        end
      end

      def import_style(source_style_id, depth:, visiting:)
        return @style_id_map[source_style_id] if @style_id_map.key?(source_style_id)

        source_node = @source_index[source_style_id]
        unless source_node
          @style_id_map[source_style_id] = source_style_id
          return source_style_id
        end

        if depth >= MAX_IMPORT_DEPTH || visiting.include?(source_style_id)
          @style_id_map[source_style_id] = source_style_id
          return source_style_id
        end

        visiting = visiting + [source_style_id]
        import_dependencies(source_node, depth: depth, visiting: visiting)

        target_node = @target_index[source_style_id]
        if target_node
          if equivalent_style?(source_node, target_node)
            @style_id_map[source_style_id] = source_style_id
            return source_style_id
          end

          new_style_id = allocate_renamed_style_id(source_style_id)
          append_imported_style(source_node, new_style_id)
          @style_id_map[source_style_id] = new_style_id
          new_style_id
        else
          append_imported_style(source_node, source_style_id)
          @style_id_map[source_style_id] = source_style_id
          source_style_id
        end
      end

      def import_dependencies(source_node, depth:, visiting:)
        dependency_style_ids(source_node).each do |dependency_id|
          import_style(dependency_id, depth: depth + 1, visiting: visiting)
        end
      end

      def dependency_style_ids(source_node)
        DEPENDENCY_ELEMENTS.filter_map do |element_name|
          source_node.at_xpath("w:#{element_name}", XML_NAMESPACES)&.[]('w:val')
        end
      end

      def append_imported_style(source_node, target_style_id)
        imported_node = import_node_into_target_styles(source_node)
        imported_node['w:styleId'] = target_style_id
        rewrite_style_dependencies(imported_node)
        append_style_to_target(imported_node)
        @target_index[target_style_id] = imported_node
      end

      def import_node_into_target_styles(source_node)
        source_node.dup(1)
      end

      def rewrite_style_dependencies(style_node)
        DEPENDENCY_ELEMENTS.each do |element_name|
          dependency_node = style_node.at_xpath("w:#{element_name}", XML_NAMESPACES)
          next unless dependency_node

          dependency_id = dependency_node['w:val']
          next unless dependency_id

          mapped_id = @style_id_map.fetch(dependency_id, dependency_id)
          dependency_node['w:val'] = mapped_id
        end
      end

      def append_style_to_target(style_node)
        styles_root = @target_doc.styles.at_xpath('//w:styles', XML_NAMESPACES)
        last_style = styles_root.xpath('./w:style', XML_NAMESPACES).last

        if last_style
          last_style.add_next_sibling(style_node)
        elsif (latent_styles = styles_root.at_xpath('./w:latentStyles', XML_NAMESPACES))
          latent_styles.add_next_sibling(style_node)
        elsif (doc_defaults = styles_root.at_xpath('./w:docDefaults', XML_NAMESPACES))
          doc_defaults.add_next_sibling(style_node)
        else
          styles_root.add_child(style_node)
        end
      end

      def allocate_renamed_style_id(original_style_id)
        loop do
          @rename_sequence += 1
          candidate = "m#{@rename_sequence}_#{original_style_id}"
          return candidate unless @target_index.key?(candidate)
        end
      end

      def equivalent_style?(source_node, target_node)
        normalize_style_xml(source_node) == normalize_style_xml(target_node)
      end

      def normalize_style_xml(node)
        node.canonicalize(Nokogiri::XML::Node::SaveOptions::AS_XML)
      rescue StandardError
        strip_whitespace_text_nodes(node.dup(1)).to_xml(save_with: Nokogiri::XML::Node::SaveOptions::AS_XML)
      end

      def strip_whitespace_text_nodes(node)
        node.children.each do |child|
          if child.text? && child.content.strip.empty?
            child.remove
          elsif child.element?
            strip_whitespace_text_nodes(child)
          end
        end
        node
      end
    end
  end
end
