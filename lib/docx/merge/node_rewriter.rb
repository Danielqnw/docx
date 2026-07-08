# frozen_string_literal: true

require 'nokogiri'

module Docx
  module Merge
    class NodeRewriter
      STYLE_ELEMENTS = %w[pStyle rStyle tblStyle].freeze
      XML_NAMESPACES = Docx::Document::XML_NAMESPACES

      def initialize(style_id_map: {}, num_id_map: {}, rid_map: {}, bookmark_id_offset: 0)
        @style_id_map = style_id_map
        @num_id_map = num_id_map
        @rid_map = rid_map
        @bookmark_id_offset = bookmark_id_offset
      end

      def rewrite(node)
        rewrite_style_references(node)
        rewrite_num_id_references(node)
        rewrite_rid_references(node)
        rewrite_bookmark_ids(node)
        node
      end

      private

      RID_NS = { 'r' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships' }.freeze
      RID_ATTRIBUTE_XPATH = './/@r:embed | .//@r:id | .//@r:link | self::*/@r:embed | self::*/@r:id | self::*/@r:link'.freeze

      def rewrite_style_references(node)
        STYLE_ELEMENTS.each do |element_name|
          node.xpath("(self::w:#{element_name} | .//w:#{element_name})", XML_NAMESPACES).each do |style_node|
            rewrite_val_from_map(style_node, @style_id_map)
          end
        end
      end

      def rewrite_num_id_references(node)
        node.xpath('(.//w:numPr | self::w:numPr)/w:numId', XML_NAMESPACES).each do |num_id_node|
          rewrite_val_from_map(num_id_node, @num_id_map)
        end
      end

      def rewrite_val_from_map(element_node, map)
        val = element_node['w:val']
        return unless val && map.key?(val)

        element_node['w:val'] = map[val]
      end

      def rewrite_rid_references(node)
        node.xpath(RID_ATTRIBUTE_XPATH, RID_NS).each do |attr|
          next unless @rid_map.key?(attr.value)

          attr.value = @rid_map[attr.value]
        end
      end

      def rewrite_bookmark_ids(node)
        return if @bookmark_id_offset.zero?

        node.xpath(
          'self::w:bookmarkStart | .//w:bookmarkStart | self::w:bookmarkEnd | .//w:bookmarkEnd',
          XML_NAMESPACES
        ).each do |bookmark_node|
          id_val = bookmark_node['w:id']
          next unless id_val&.match?(/\A\d+\z/)

          bookmark_node['w:id'] = (id_val.to_i + @bookmark_id_offset).to_s
        end
      end
    end
  end
end
