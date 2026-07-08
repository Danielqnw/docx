# frozen_string_literal: true

require 'nokogiri'

module Docx
  module Merge
    class NumberingImporter
      STYLE_LINK_ELEMENTS = %w[styleLink numStyleLink].freeze
      XML_NAMESPACES = Docx::Document::XML_NAMESPACES

      attr_reader :num_id_map, :abstract_num_id_map

      def initialize(target_doc, source_doc, style_id_map: {})
        @target_doc = target_doc
        @source_doc = source_doc
        @style_id_map = style_id_map
        @num_id_map = {}
        @abstract_num_id_map = {}
        @target_numbering = nil
        @target_index = nil
        @source_index = nil
      end

      def import(source_num_id)
        source_num_id = source_num_id.to_s
        return source_num_id if @source_doc.numbering.nil?

        ensure_target_ready!
        return @num_id_map[source_num_id] if @num_id_map.key?(source_num_id)

        source_num_node = source_index[:num][source_num_id]
        unless source_num_node
          @num_id_map[source_num_id] = source_num_id
          return source_num_id
        end

        abstract_num_id_ref = source_num_node.at_xpath('./w:abstractNumId', XML_NAMESPACES)&.[]('w:val')
        import_abstract_num(abstract_num_id_ref) if abstract_num_id_ref

        import_num(source_num_id, source_num_node)
      end

      def import_all
        return if @source_doc.numbering.nil?

        ensure_target_ready!
        source_index[:num].each_key do |num_id|
          import(num_id)
        end
      end

      private

      def ensure_target_ready!
        @target_numbering = @target_doc.ensure_numbering!
        @target_index = build_index(@target_numbering) if @target_index.nil?
      end

      def source_index
        @source_index ||= build_index(@source_doc.numbering)
      end

      def build_index(numbering_doc)
        return { abstract_num: {}, num: {} } unless numbering_doc

        abstract_nums = numbering_doc.xpath('//w:numbering/w:abstractNum', XML_NAMESPACES).each_with_object({}) do |node, index|
          abs_id = node['w:abstractNumId']
          index[abs_id] = node if abs_id
        end

        nums = numbering_doc.xpath('//w:numbering/w:num', XML_NAMESPACES).each_with_object({}) do |node, index|
          num_id = node['w:numId']
          index[num_id] = node if num_id
        end

        { abstract_num: abstract_nums, num: nums }
      end

      def next_abstract_num_id
        ids = @target_index[:abstract_num].keys.map(&:to_i)
        base = ids.empty? ? -1 : ids.max
        (base + 1).to_s
      end

      def next_num_id
        ids = @target_index[:num].keys.map(&:to_i)
        base = ids.empty? ? 0 : ids.max
        (base + 1).to_s
      end

      def import_abstract_num(source_abs_id)
        source_abs_id = source_abs_id.to_s
        return @abstract_num_id_map[source_abs_id] if @abstract_num_id_map.key?(source_abs_id)

        source_abs_node = source_index[:abstract_num][source_abs_id]
        unless source_abs_node
          @abstract_num_id_map[source_abs_id] = source_abs_id
          return source_abs_id
        end

        new_abs_id = next_abstract_num_id
        imported_node = source_abs_node.dup(1)
        imported_node['w:abstractNumId'] = new_abs_id
        rewrite_style_links(imported_node)
        append_abstract_num(imported_node)

        @abstract_num_id_map[source_abs_id] = new_abs_id
        @target_index[:abstract_num][new_abs_id] = imported_node
        new_abs_id
      end

      def import_num(source_num_id, source_num_node)
        source_num_id = source_num_id.to_s
        return @num_id_map[source_num_id] if @num_id_map.key?(source_num_id)

        new_num_id = next_num_id
        imported_node = source_num_node.dup(1)
        imported_node['w:numId'] = new_num_id

        abstract_num_id_ref = source_num_node.at_xpath('./w:abstractNumId', XML_NAMESPACES)&.[]('w:val')
        if abstract_num_id_ref
          mapped_abs_id = @abstract_num_id_map.fetch(abstract_num_id_ref.to_s, abstract_num_id_ref.to_s)
          abs_ref_node = imported_node.at_xpath('./w:abstractNumId', XML_NAMESPACES)
          abs_ref_node['w:val'] = mapped_abs_id if abs_ref_node
        end

        append_num(imported_node)

        @num_id_map[source_num_id] = new_num_id
        @target_index[:num][new_num_id] = imported_node
        new_num_id
      end

      def rewrite_style_links(abstract_num_node)
        STYLE_LINK_ELEMENTS.each do |element_name|
          link_node = abstract_num_node.at_xpath("w:#{element_name}", XML_NAMESPACES)
          next unless link_node

          style_id = link_node['w:val']
          next unless style_id

          link_node['w:val'] = @style_id_map.fetch(style_id, style_id)
        end
      end

      def numbering_root
        @target_numbering.at_xpath('//w:numbering', XML_NAMESPACES)
      end

      def append_abstract_num(node)
        root = numbering_root
        last_abs = root.xpath('./w:abstractNum', XML_NAMESPACES).last

        if last_abs
          last_abs.add_next_sibling(node)
        elsif (first_num = root.xpath('./w:num', XML_NAMESPACES).first)
          first_num.add_previous_sibling(node)
        else
          root.add_child(node)
        end
      end

      def append_num(node)
        numbering_root.add_child(node)
      end
    end
  end
end
