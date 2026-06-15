# frozen_string_literal: true

require 'spec_helper'
require 'docx/document'
require 'tempfile'

describe 'table merge write API' do
  fixtures_path = File.join('spec/fixtures/tables')

  before(:all) do
    TableFixtureBuilder.write_all_fixtures(fixtures_path)
  end

  def open_plain_table
    doc = Docx::Document.open(File.join('spec/fixtures/tables', 'plain_3x3.docx'))
    doc.tables.first
  end

  def physical_cell_count(table, row)
    table.rows[row].cells.size
  end

  def physical_tc_at_col0(table, row)
    table.rows[row].cells.first.node
  end

  def grid_span_val(tc_node)
    val = tc_node.at_xpath('w:tcPr/w:gridSpan/@w:val')&.value
    val ? val.to_i : 1
  end

  def vmerge_mode(tc_node)
    vmerge = tc_node.at_xpath('w:tcPr/w:vMerge')
    return nil unless vmerge

    val = vmerge['w:val'] || vmerge['val'] || vmerge.at_xpath('@w:val')&.value
    val == 'restart' ? :restart : :continue
  end

  def round_trip_without_stderr(doc)
    temp_file = Tempfile.new(['table_merge_write', '.docx'])
    temp_path = temp_file.path
    temp_file.close

    begin
      doc.save(temp_path)

      reopened = nil
      expect do
        reopened = Docx::Document.open(temp_path)
      end.to output('').to_stderr

      reopened
    ensure
      File.delete(temp_path) if temp_path && File.exist?(temp_path)
    end
  end

  describe 'horizontal merge' do
    let(:fixture_path) { File.join(fixtures_path, 'plain_3x3.docx') }

    it 'merges two adjacent cells in one row and round-trips cleanly' do
      doc = Docx::Document.open(fixture_path)
      table = doc.tables.first

      expect(table.row_count).to eq(3)
      expect(table.column_count).to eq(3)
      expect(physical_cell_count(table, 0)).to eq(3)

      table.merge_cells(0, 0, 0, 1)

      anchor = table.cell_at(0, 0)
      expect(anchor.colspan).to eq(2)
      expect(anchor.merge_anchor?).to eq(true)
      expect(anchor.text).to eq('r0c0')
      expect(physical_cell_count(table, 0)).to eq(2)
      expect(table.row_count).to eq(3)
      expect(table.column_count).to eq(3)
      expect(table.cell_at(0, 2).text).to eq('r0c2')
      expect(table.cell_at(1, 0).text).to eq('r1c0')

      reopened = round_trip_without_stderr(doc)
      reopened_table = reopened.tables.first

      expect(reopened_table.row_count).to eq(3)
      expect(reopened_table.column_count).to eq(3)
      expect(reopened_table.cell_at(0, 0).colspan).to eq(2)
      expect(reopened_table.cell_at(0, 0).text).to eq('r0c0')
      expect(reopened_table.cell_at(0, 2).text).to eq('r0c2')
      expect(reopened_table.cell_at(1, 0).text).to eq('r1c0')
    end

    it 'is a no-op for a single-cell range' do
      table = open_plain_table

      table.merge_cells(1, 1, 1, 1)

      cell = table.cell_at(1, 1)
      expect(cell.colspan).to eq(1)
      expect(physical_cell_count(table, 1)).to eq(3)
      expect(table.row_count).to eq(3)
      expect(table.column_count).to eq(3)
    end

    it 'raises InvalidMergeRange for out-of-bounds column' do
      table = open_plain_table

      expect do
        table.merge_cells(0, 0, 0, 3)
      end.to raise_error(Docx::Errors::InvalidMergeRange)
    end

    it 'raises InvalidMergeRange when col0 is greater than col1' do
      table = open_plain_table

      expect do
        table.merge_cells(0, 2, 0, 1)
      end.to raise_error(Docx::Errors::InvalidMergeRange)
    end
  end

  describe 'vertical and rectangular merge' do
    it 'merges a 2x2 rectangle and round-trips cleanly' do
      doc = Docx::Document.open(File.join(fixtures_path, 'plain_3x3.docx'))
      table = doc.tables.first

      table.merge_cells(0, 0, 1, 1)

      anchor = table.cell_at(0, 0)
      expect(anchor.colspan).to eq(2)
      expect(anchor.rowspan).to eq(2)
      expect(anchor.merge_anchor?).to eq(true)
      expect(anchor.text).to eq('r0c0')
      expect(table.merged?(0, 0)).to eq(true)
      expect(table.merged?(0, 1)).to eq(true)
      expect(table.merged?(1, 0)).to eq(true)
      expect(table.merged?(1, 1)).to eq(true)
      expect(table.merged?(2, 2)).to eq(false)
      expect(table.column_count).to eq(3)
      expect(table.row_count).to eq(3)
      expect(physical_cell_count(table, 2)).to eq(3)
      expect(physical_cell_count(table, 0)).to eq(2)
      expect(physical_cell_count(table, 1)).to eq(2)

      row0_tc = physical_tc_at_col0(table, 0)
      row1_tc = physical_tc_at_col0(table, 1)
      expect(grid_span_val(row0_tc)).to eq(2)
      expect(grid_span_val(row1_tc)).to eq(2)
      expect(vmerge_mode(row0_tc)).to eq(:restart)
      expect(vmerge_mode(row1_tc)).to eq(:continue)

      reopened = round_trip_without_stderr(doc)
      reopened_table = reopened.tables.first

      reopened_anchor = reopened_table.cell_at(0, 0)
      expect(reopened_anchor.colspan).to eq(2)
      expect(reopened_anchor.rowspan).to eq(2)
      expect(reopened_anchor.merge_anchor?).to eq(true)
      expect(reopened_anchor.text).to eq('r0c0')
      expect(reopened_table.merged?(0, 0)).to eq(true)
      expect(reopened_table.merged?(0, 1)).to eq(true)
      expect(reopened_table.merged?(1, 0)).to eq(true)
      expect(reopened_table.merged?(1, 1)).to eq(true)
      expect(reopened_table.merged?(2, 2)).to eq(false)
      expect(reopened_table.column_count).to eq(3)
      expect(reopened_table.row_count).to eq(3)
      expect(physical_cell_count(reopened_table, 2)).to eq(3)

      reopened_row0_tc = physical_tc_at_col0(reopened_table, 0)
      reopened_row1_tc = physical_tc_at_col0(reopened_table, 1)
      expect(grid_span_val(reopened_row0_tc)).to eq(2)
      expect(grid_span_val(reopened_row1_tc)).to eq(2)
      expect(vmerge_mode(reopened_row0_tc)).to eq(:restart)
      expect(vmerge_mode(reopened_row1_tc)).to eq(:continue)
    end

    it 'merges three rows in one column and round-trips cleanly' do
      doc = Docx::Document.open(File.join(fixtures_path, 'plain_3x3.docx'))
      table = doc.tables.first

      table.merge_cells(0, 0, 2, 0)

      anchor = table.cell_at(0, 0)
      expect(anchor.rowspan).to eq(3)
      expect(anchor.colspan).to eq(1)
      expect(anchor.merge_anchor?).to eq(true)
      expect(anchor.text).to eq('r0c0')
      expect(table.merged?(0, 0)).to eq(true)
      expect(table.merged?(1, 0)).to eq(true)
      expect(table.merged?(2, 0)).to eq(true)

      continuation_row1 = table.rows[1].cells[0]
      continuation_row2 = table.rows[2].cells[0]
      expect(continuation_row1.merge_continuation?).to eq(true)
      expect(continuation_row2.merge_continuation?).to eq(true)

      reopened = round_trip_without_stderr(doc)
      reopened_table = reopened.tables.first

      reopened_anchor = reopened_table.cell_at(0, 0)
      expect(reopened_anchor.rowspan).to eq(3)
      expect(reopened_anchor.colspan).to eq(1)
      expect(reopened_anchor.merge_anchor?).to eq(true)
      expect(reopened_anchor.text).to eq('r0c0')
      expect(reopened_table.merged?(0, 0)).to eq(true)
      expect(reopened_table.merged?(1, 0)).to eq(true)
      expect(reopened_table.merged?(2, 0)).to eq(true)
      expect(reopened_table.rows[1].cells[0].merge_continuation?).to eq(true)
      expect(reopened_table.rows[2].cells[0].merge_continuation?).to eq(true)
    end

    it 'merges three columns in one row and round-trips cleanly' do
      doc = Docx::Document.open(File.join(fixtures_path, 'plain_3x3.docx'))
      table = doc.tables.first

      table.merge_cells(0, 0, 0, 2)

      anchor = table.cell_at(0, 0)
      expect(anchor.colspan).to eq(3)
      expect(physical_cell_count(table, 0)).to eq(1)
      expect(table.column_count).to eq(3)

      reopened = round_trip_without_stderr(doc)
      reopened_table = reopened.tables.first

      expect(reopened_table.cell_at(0, 0).colspan).to eq(3)
      expect(physical_cell_count(reopened_table, 0)).to eq(1)
      expect(reopened_table.column_count).to eq(3)
    end

    it 'raises MergeConflict when the range overlaps an existing merge' do
      table = open_plain_table

      table.merge_cells(0, 0, 1, 1)

      expect do
        table.merge_cells(1, 1, 2, 2)
      end.to raise_error(Docx::Errors::MergeConflict)
    end
  end
end
