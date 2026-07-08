# frozen_string_literal: true

require 'nokogiri'
require 'set'

module Docx
  module Merge
    class Importer
      STYLE_ELEMENTS = %w[pStyle rStyle tblStyle].freeze
      XML_NAMESPACES = Docx::Document::XML_NAMESPACES

      def initialize(target_doc, source_doc)
        @target_doc = target_doc
        @source_doc = source_doc
        @styles_importer = StylesImporter.new(target_doc, source_doc)
        @numbering_importer = NumberingImporter.new(
          target_doc,
          source_doc,
          style_id_map: @styles_importer.style_id_map
        )
      end

      def import(node)
        collect_style_ids(node).each do |style_id|
          @styles_importer.import(style_id)
        end

        collect_num_ids(node).each do |num_id|
          @numbering_importer.import(num_id)
        end

        imported = node.dup(1)
        NodeRewriter.new(
          style_id_map: @styles_importer.style_id_map,
          num_id_map: @numbering_importer.num_id_map
        ).rewrite(imported)
      end

      def style_id_map
        @styles_importer.style_id_map
      end

      def num_id_map
        @numbering_importer.num_id_map
      end

      def abstract_num_id_map
        @numbering_importer.abstract_num_id_map
      end

      private

      def collect_style_ids(node)
        ids = Set.new
        STYLE_ELEMENTS.each do |element_name|
          node.xpath("(self::w:#{element_name} | .//w:#{element_name})/@w:val", XML_NAMESPACES).each do |attr|
            ids.add(attr.value)
          end
        end
        ids.to_a
      end

      def collect_num_ids(node)
        node.xpath('(self::w:numPr | .//w:numPr)/w:numId/@w:val', XML_NAMESPACES).map(&:value).uniq
      end
    end
  end
end
