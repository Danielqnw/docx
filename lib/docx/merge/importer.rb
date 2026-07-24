# frozen_string_literal: true

require 'nokogiri'
require 'set'

module Docx
  module Merge
    class Importer
      STYLE_ELEMENTS = %w[pStyle rStyle tblStyle].freeze
      XML_NAMESPACES = Docx::Document::XML_NAMESPACES
      RID_NS = { 'r' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships' }.freeze
      RID_ATTRIBUTE_XPATH = './/@r:embed | .//@r:id | .//@r:link | self::*/@r:embed | self::*/@r:id | self::*/@r:link'.freeze

      def initialize(target_doc, source_doc)
        @target_doc = target_doc
        @source_doc = source_doc
        @bookmark_id_offset = compute_bookmark_id_offset
        @styles_importer = StylesImporter.new(target_doc, source_doc)
        @numbering_importer = NumberingImporter.new(
          target_doc,
          source_doc,
          styles_importer: @styles_importer
        )
        @media_importer = MediaImporter.new(target_doc, source_doc)
      end

      attr_reader :bookmark_id_offset

      def import(node)
        collect_style_ids(node).each do |style_id|
          @styles_importer.import(style_id)
        end

        collect_num_ids(node).each do |num_id|
          @numbering_importer.import(num_id)
        end

        collect_relationship_ids(node).each do |rid|
          @media_importer.import(rid)
        end

        imported = node.dup(1)
        NodeRewriter.new(
          style_id_map: @styles_importer.style_id_map,
          num_id_map: @numbering_importer.num_id_map,
          rid_map: @media_importer.rid_map,
          bookmark_id_offset: @bookmark_id_offset
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

      def rid_map
        @media_importer.rid_map
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

      def collect_relationship_ids(node)
        node.xpath(RID_ATTRIBUTE_XPATH, RID_NS).map(&:value).uniq
      end

      def compute_bookmark_id_offset
        ids = @target_doc.doc.xpath(
          '//w:bookmarkStart/@w:id | //w:bookmarkEnd/@w:id',
          XML_NAMESPACES
        ).map(&:value).select { |val| val.match?(/\A\d+\z/) }.map(&:to_i)
        ids.empty? ? 0 : ids.max + 1
      end
    end
  end
end
