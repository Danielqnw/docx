require 'docx/errors'
require 'docx/containers/table_cell'

module Docx
  module Elements
    module Containers
      module TableMerge
        def merge_cells(row0, col0, row1, col1)
          validate_merge_range!(row0, col0, row1, col1)

          return if row0 == row1 && col0 == col1

          detect_merge_overlap!(row0, col0, row1, col1)

          if row0 == row1
            merge_cells_horizontal(row0, col0, col1)
          else
            merge_cells_rectangular(row0, col0, row1, col1)
          end
        end

        def unmerge_cells(row, col)
          slot = grid.slots[row][col]
          raise Docx::Errors::InvalidMergeTarget if slot.nil?

          anchor = slot.anchor
          if anchor.row != row || anchor.col != col
            raise Docx::Errors::InvalidMergeTarget
          end

          return if anchor.colspan == 1 && anchor.rowspan == 1

          r0 = anchor.row
          c0 = anchor.col
          height = anchor.rowspan
          width = anchor.colspan

          (r0...(r0 + height)).each do |r|
            tc = physical_tc_at(r, c0)

            if width > 1
              remove_grid_span(tc)
              insert_blank_cells_after(tc, width - 1)
            end

            remove_vmerge(tc)
          end

          invalidate_grid!
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

        def detect_merge_overlap!(row0, col0, row1, col1)
          (row0..row1).each do |row|
            (col0..col1).each do |col|
              slot = grid.slots[row][col]
              next if slot.nil?

              anchor = slot.anchor
              next unless anchor.colspan > 1 || anchor.rowspan > 1

              raise Docx::Errors::MergeConflict,
                    "merge range (#{row0},#{col0})..(#{row1},#{col1}) overlaps existing merge"
            end
          end
        end

        def merge_cells_horizontal(row, col0, col1)
          apply_horizontal_span_in_row(row, col0, col1)
          invalidate_grid!
        end

        def merge_cells_rectangular(row0, col0, row1, col1)
          (row0..row1).each do |row|
            apply_horizontal_span_in_row(row, col0, col1) if col1 > col0
          end

          invalidate_grid!

          (row0..row1).each do |row|
            ensure_continuation_cell!(row, col0)
            tc_node = physical_tc_at(row, col0)

            if row == row0
              set_vmerge(tc_node, :restart)
            else
              set_vmerge(tc_node, :continue)
              Containers::TableCell.new(tc_node).blank!
            end
          end

          invalidate_grid!
        end

        def apply_horizontal_span_in_row(row, col0, col1)
          logical_width = col1 - col0 + 1
          tc_nodes = physical_tc_nodes_in_row_range(row, col0, col1)

          anchor_node = tc_nodes.first
          set_grid_span(anchor_node, logical_width)

          tc_nodes.drop(1).each(&:remove)
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

        def physical_tc_at(row, col)
          col_pos = 0
          tr_node = @node.xpath('w:tr')[row]

          tr_node.xpath('w:tc').each do |tc|
            return tc if col_pos == col

            col_pos += read_grid_span(tc)
          end

          nil
        end

        def ensure_continuation_cell!(row, col)
          return if physical_tc_at(row, col)

          tr_node = @node.xpath('w:tr')[row]
          insert_index = physical_insert_index_for_col(row, col)

          new_tc = Nokogiri::XML::Node.new('w:tc', tr_node.document)
          tc_pr = Nokogiri::XML::Node.new('w:tcPr', tr_node.document)
          paragraph = Nokogiri::XML::Node.new('w:p', tr_node.document)
          new_tc.add_child(tc_pr)
          new_tc.add_child(paragraph)

          existing_tcs = tr_node.xpath('w:tc')
          if insert_index >= existing_tcs.length
            tr_node.add_child(new_tc)
          else
            existing_tcs[insert_index].add_previous_sibling(new_tc)
          end
        end

        def physical_insert_index_for_col(row, target_col)
          col = 0
          tr_node = @node.xpath('w:tr')[row]

          tr_node.xpath('w:tc').each_with_index do |tc, idx|
            return idx if col == target_col

            col += read_grid_span(tc)
          end

          tr_node.xpath('w:tc').length
        end

        def read_grid_span(tc_node)
          val = tc_node.at_xpath('w:tcPr/w:gridSpan/@w:val')&.value
          val ? val.to_i : 1
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

        def set_vmerge(tc_node, mode)
          tc_pr = ensure_tc_pr!(tc_node)
          vmerge = tc_pr.at_xpath('w:vMerge')

          if vmerge
            update_vmerge_node(vmerge, mode)
          else
            vmerge = Nokogiri::XML::Node.new('w:vMerge', tc_node.document)
            update_vmerge_node(vmerge, mode)
            tc_pr.add_child(vmerge)
          end
        end

        def update_vmerge_node(vmerge_node, mode)
          if mode == :restart
            set_w_val(vmerge_node, 'restart')
          else
            vmerge_node.remove_attribute('w:val')
          end
        end

        def set_w_val(node, value)
          node['w:val'] = value.to_s
        end

        def remove_grid_span(tc)
          tc.at_xpath('w:tcPr/w:gridSpan')&.remove
        end

        def remove_vmerge(tc)
          tc.at_xpath('w:tcPr/w:vMerge')&.remove
        end

        def insert_blank_cells_after(tc, count)
          insert_after = tc

          count.times do
            new_tc = blank_tc_node(tc.document)
            insert_after.add_next_sibling(new_tc)
            insert_after = new_tc
          end
        end

        def blank_tc_node(document)
          new_tc = Nokogiri::XML::Node.new('w:tc', document)
          tc_pr = Nokogiri::XML::Node.new('w:tcPr', document)
          paragraph = Nokogiri::XML::Node.new('w:p', document)
          new_tc.add_child(tc_pr)
          new_tc.add_child(paragraph)
          new_tc
        end
      end
    end
  end
end
