require 'docx/errors'

module Docx
  module Elements
    module Containers
      module TableMerge
        def merge_cells(row0, col0, row1, col1)
          validate_merge_range!(row0, col0, row1, col1)

          return if row0 == row1 && col0 == col1

          if row0 != row1
            raise NotImplementedError, 'vertical and rectangular merge not yet implemented'
          end

          merge_cells_horizontal(row0, col0, col1)
        end

        private

        def validate_merge_range!(row0, col0, row1, col1)
          if row0 > row1 || col0 > col1 ||
             row0.negative? || col0.negative? ||
             row1 >= row_count || col1 >= column_count
            raise Docx::Errors::InvalidMergeRange,
                  "invalid merge range (#{row0},#{col0})..(#{row1},#{col1})"
          end
        end

        def merge_cells_horizontal(row, col0, col1)
          logical_width = col1 - col0 + 1
          tc_nodes = physical_tc_nodes_in_row_range(row, col0, col1)

          anchor_node = tc_nodes.first
          set_grid_span(anchor_node, logical_width)

          tc_nodes.drop(1).each(&:remove)

          invalidate_grid!
        end

        def physical_tc_nodes_in_row_range(row, col0, col1)
          nodes = []
          seen = {}

          (col0..col1).each do |col|
            slot = grid.slots[row][col]
            node = slot.anchor.node
            next if seen[node.object_id]

            seen[node.object_id] = true
            nodes << node
          end

          nodes
        end

        def ensure_tc_pr!(tc_node)
          tc_pr = tc_node.at_xpath('w:tcPr')
          return tc_pr if tc_pr

          tc_pr = Nokogiri::XML::Node.new('w:tcPr', tc_node.document)
          tc_node.prepend_child(tc_pr)
          tc_pr
        end

        def set_grid_span(tc_node, n)
          tc_pr = ensure_tc_pr!(tc_node)
          grid_span = tc_pr.at_xpath('w:gridSpan')

          if grid_span
            set_w_val(grid_span, n)
          else
            grid_span = Nokogiri::XML::Node.new('w:gridSpan', tc_node.document)
            set_w_val(grid_span, n)
            tc_pr.add_child(grid_span)
          end
        end

        def set_w_val(node, value)
          node['w:val'] = value.to_s
        end
      end
    end
  end
end
