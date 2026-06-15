require 'docx/errors'
require 'docx/containers/text_run'
require 'docx/containers/container'
require 'docx/containers/table_grid'

module Docx
  module Elements
    module Containers
      class TableCell
        include Container
        include Elements::Element

        def self.tag
          'tc'
        end

        def initialize(node, grid_slot: nil, logical_row: nil, logical_col: nil)
          @node = node
          @properties_tag = 'tcPr'
          @grid_slot = grid_slot
          @logical_row = logical_row
          @logical_col = logical_col
        end

        # Return text of paragraph's cell
        def to_s
          paragraphs.map(&:text).join('')
        end

        # Array of paragraphs contained within cell
        def paragraphs
          @node.xpath('w:p').map {|p_node| Containers::Paragraph.new(p_node) }
        end

        # Iterate over each text run within a paragraph's cell
        def each_paragraph
          paragraphs.each { |tr| yield(tr) }
        end

        def colspan
          grid_span_value
        end

        def rowspan
          return 1 if merge_continuation?
          return 1 unless vmerge_state == :restart

          slot = resolved_grid_slot
          slot ? slot.rowspan : 1
        end

        def merged?
          colspan > 1 || !vmerge_state.nil?
        end

        def merge_anchor?
          !merge_continuation? && (colspan > 1 || vmerge_state == :restart)
        end

        def merge_continuation?
          vmerge_state == :continue
        end

        def unmerge!
          unless merge_anchor?
            raise Docx::Errors::InvalidMergeTarget
          end

          slot = resolved_grid_slot
          raise Docx::Errors::InvalidMergeTarget unless slot

          if !@logical_row.nil? &&
             (@logical_row != slot.row || @logical_col != slot.col)
            raise Docx::Errors::InvalidMergeTarget
          end

          table_node = @node.at_xpath('ancestor::w:tbl')
          raise Docx::Errors::InvalidMergeTarget unless table_node

          require 'docx/containers/table'
          table = Containers::Table.new(table_node)
          table.unmerge_cells(slot.row, slot.col)
        end

        alias_method :text, :to_s

        private

        def grid_span_value
          val = @node.at_xpath('w:tcPr/w:gridSpan/@w:val')&.value
          val ? val.to_i : 1
        end

        def vmerge_state
          @vmerge_state ||= begin
            vmerge_node = @node.at_xpath('w:tcPr/w:vMerge')
            unless vmerge_node
              nil
            else
              val = vmerge_node['val'] || vmerge_node.at_xpath('@w:val')&.value
              val == 'restart' ? :restart : :continue
            end
          end
        end

        def resolved_grid_slot
          if @grid_slot && @grid_slot.node == @node
            return @grid_slot
          end

          @resolved_grid_slot ||= anchor_slot_from_grid
        end

        def anchor_slot_from_grid
          table_node = @node.at_xpath('ancestor::w:tbl')
          return nil unless table_node

          grid = Containers::TableGrid.new(table_node)
          grid.each_anchor.find { |anchor| anchor.node == @node }
        end
      end
    end
  end
end
