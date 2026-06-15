require 'docx/containers/table_cell'

module Docx
  module Elements
    module Containers
      CellSlot = Struct.new(
        :node, :row, :col, :colspan, :rowspan,
        :anchor, :continuation,
        keyword_init: true
      )

      class TableGrid
        attr_reader :slots, :row_count, :column_count

        def initialize(table_node)
          @node = table_node
          build_slots
        end

        def cell_at(row, col)
          return nil if row.negative? || col.negative?
          return nil if row >= row_count || col >= column_count

          slot = @slots[row][col]
          return nil if slot.nil?

          slot.anchor
        end

        def each_anchor
          return enum_for(:each_anchor) unless block_given?

          seen = {}
          @slots.each do |row|
            row.each do |slot|
              next if slot.nil?

              anchor = slot.anchor
              next if seen[anchor.object_id]

              seen[anchor.object_id] = true
              yield(anchor)
            end
          end
        end

        private

        def build_slots
          row_nodes = @node.xpath('w:tr')
          @row_count = row_nodes.count
          row_layouts = row_nodes.map { |tr| parse_row(tr) }

          tbl_grid_count = @node.xpath('w:tblGrid/w:gridCol').count
          row_widths = row_layouts.map { |cells| cells.sum { |c| c[:colspan] } }
          max_row_width = row_widths.max || 0
          @column_count = [max_row_width, tbl_grid_count].max

          if tbl_grid_count.positive? && row_widths.any? { |w| w.positive? && w != tbl_grid_count }
            warn(
              "Table grid width mismatch: tblGrid has #{tbl_grid_count} columns " \
              "but row gridSpan sums are #{row_widths.inspect}"
            )
          end

          @slots = Array.new(@row_count) { Array.new(@column_count) }

          row_layouts.each_with_index do |cells, row|
            col = 0
            cells.each do |cell_info|
              colspan = cell_info[:colspan]
              vmerge = cell_info[:vmerge]

              if vmerge == :continue
                anchor_slot = @slots[row - 1][col]
                anchor_slot.rowspan += 1
                colspan.times { |offset| @slots[row][col + offset] = anchor_slot }
              else
                anchor_slot = CellSlot.new(
                  node: cell_info[:node],
                  row: row,
                  col: col,
                  colspan: colspan,
                  rowspan: 1,
                  anchor: nil,
                  continuation: false
                )
                anchor_slot.anchor = anchor_slot
                colspan.times { |offset| @slots[row][col + offset] = anchor_slot }
              end

              col += colspan
            end
          end
        end

        def parse_row(tr_node)
          tr_node.xpath('w:tc').map do |tc|
            {
              node: tc,
              colspan: read_grid_span(tc),
              vmerge: read_vmerge(tc)
            }
          end
        end

        def read_grid_span(tc)
          val = tc.at_xpath('w:tcPr/w:gridSpan/@w:val')&.value
          val ? val.to_i : 1
        end

        def read_vmerge(tc)
          vmerge_node = tc.at_xpath('w:tcPr/w:vMerge')
          return nil unless vmerge_node

          val = vmerge_node['val'] || vmerge_node.at_xpath('@w:val')&.value
          val == 'restart' ? :restart : :continue
        end
      end
    end
  end
end
