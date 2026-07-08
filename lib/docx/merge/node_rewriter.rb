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
        node
      end

      private

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
    end
  end
end
