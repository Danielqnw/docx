# frozen_string_literal: true

require 'spec_helper'
require 'docx/document'
require 'tempfile'

describe 'table unmerge API' do
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

  def grid_span_val(tc_node)
    val = tc_node.at_xpath('w:tcPr/w:gridSpan/@w:val')&.value
    val ? val.to_i : 1
  end

  def grid_span_sum(table, row)
    table.rows[row].cells.sum { |cell| grid_span_val(cell.node) }
  end

  def round_trip_without_stderr(doc)
    temp_file = Tempfile.new(['table_unmerge', '.docx'])
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

  describe 'horizontal unmerge (1x3)' do
    it 'splits a merged row and round-trips cleanly' do
      doc = Docx::Document.open(File.join(fixtures_path, 'plain_3x3.docx'))
      table = doc.tables.first

      table.merge_cells(0, 0, 0, 2)
      table.unmerge_cells(0, 0)

      expect(table.cell_at(0, 0).colspan).to eq(1)
      expect(table.merged?(0, 0)).to eq(false)
      expect(table.cell_at(0, 0).text).to eq('r0c0')
      expect(physical_cell_count(table, 0)).to eq(3)
      expect(grid_span_sum(table, 0)).to eq(table.column_count)

      reopened = round_trip_without_stderr(doc)
      reopened_table = reopened.tables.first

      expect(reopened_table.row_count).to eq(3)
      expect(reopened_table.column_count).to eq(3)
      expect(physical_cell_count(reopened_table, 0)).to eq(3)
      expect(reopened_table.cell_at(0, 0).colspan).to eq(1)
      expect(reopened_table.merged?(0, 0)).to eq(false)
      expect(reopened_table.cell_at(0, 0).text).to eq('r0c0')
      expect(grid_span_sum(reopened_table, 0)).to eq(reopened_table.column_count)
    end
  end

  describe 'vertical unmerge (3x1)' do
    it 'splits a merged column and round-trips cleanly' do
      doc = Docx::Document.open(File.join(fixtures_path, 'plain_3x3.docx'))
      table = doc.tables.first

      table.merge_cells(0, 0, 2, 0)
      table.unmerge_cells(0, 0)

      expect(table.cell_at(0, 0).rowspan).to eq(1)
      expect(table.cell_at(0, 0).text).to eq('r0c0')
      expect(table.merged?(0, 0)).to eq(false)
      expect(table.merged?(1, 0)).to eq(false)
      expect(table.merged?(2, 0)).to eq(false)
      expect(table.rows[0].cells[0].merge_continuation?).to eq(false)
      expect(table.rows[1].cells[0].merge_continuation?).to eq(false)
      expect(table.rows[2].cells[0].merge_continuation?).to eq(false)

      reopened = round_trip_without_stderr(doc)
      reopened_table = reopened.tables.first

      expect(reopened_table.cell_at(0, 0).rowspan).to eq(1)
      expect(reopened_table.cell_at(0, 0).text).to eq('r0c0')
      expect(reopened_table.merged?(0, 0)).to eq(false)
      expect(reopened_table.merged?(1, 0)).to eq(false)
      expect(reopened_table.merged?(2, 0)).to eq(false)
      expect(reopened_table.rows[0].cells[0].merge_continuation?).to eq(false)
      expect(reopened_table.rows[1].cells[0].merge_continuation?).to eq(false)
      expect(reopened_table.rows[2].cells[0].merge_continuation?).to eq(false)
    end
  end

  describe 'rectangular unmerge (2x2)' do
    it 'splits a merged rectangle and round-trips cleanly' do
      doc = Docx::Document.open(File.join(fixtures_path, 'plain_3x3.docx'))
      table = doc.tables.first

      table.merge_cells(0, 0, 1, 1)
      table.unmerge_cells(0, 0)

      expect(table.cell_at(0, 0).colspan).to eq(1)
      expect(table.cell_at(0, 0).rowspan).to eq(1)
      expect(table.merged?(0, 0)).to eq(false)
      expect(table.cell_at(0, 0).text).to eq('r0c0')
      expect(physical_cell_count(table, 0)).to eq(3)
      expect(physical_cell_count(table, 1)).to eq(3)
      expect(grid_span_sum(table, 0)).to eq(table.column_count)
      expect(grid_span_sum(table, 1)).to eq(table.column_count)
      (0...table.row_count).each do |row|
        (0...table.column_count).each do |col|
          expect(table.merged?(row, col)).to eq(false)
        end
      end

      reopened = round_trip_without_stderr(doc)
      reopened_table = reopened.tables.first

      expect(reopened_table.cell_at(0, 0).colspan).to eq(1)
      expect(reopened_table.cell_at(0, 0).rowspan).to eq(1)
      expect(reopened_table.cell_at(0, 0).text).to eq('r0c0')
      expect(physical_cell_count(reopened_table, 0)).to eq(3)
      expect(physical_cell_count(reopened_table, 1)).to eq(3)
      expect(grid_span_sum(reopened_table, 0)).to eq(reopened_table.column_count)
      expect(grid_span_sum(reopened_table, 1)).to eq(reopened_table.column_count)
      (0...reopened_table.row_count).each do |row|
        (0...reopened_table.column_count).each do |col|
          expect(reopened_table.merged?(row, col)).to eq(false)
        end
      end
    end
  end

  describe 'TableCell#unmerge!' do
    it 'delegates to table.unmerge_cells from the merge anchor' do
      doc = Docx::Document.open(File.join(fixtures_path, 'plain_3x3.docx'))
      table = doc.tables.first

      table.merge_cells(0, 0, 1, 1)
      table.cell_at(0, 0).unmerge!

      expect(table.cell_at(0, 0).colspan).to eq(1)
      expect(table.cell_at(0, 0).rowspan).to eq(1)
      expect(table.merged?(0, 0)).to eq(false)
      expect(table.cell_at(0, 0).text).to eq('r0c0')
      expect(physical_cell_count(table, 0)).to eq(3)
      expect(physical_cell_count(table, 1)).to eq(3)

      reopened = round_trip_without_stderr(doc)
      reopened_table = reopened.tables.first

      expect(reopened_table.cell_at(0, 0).colspan).to eq(1)
      expect(reopened_table.cell_at(0, 0).rowspan).to eq(1)
      expect(reopened_table.merged?(0, 0)).to eq(false)
      expect(reopened_table.cell_at(0, 0).text).to eq('r0c0')
    end
  end

  describe 'invalid unmerge targets' do
    it 'raises InvalidMergeTarget for non-anchor coordinates' do
      table = open_plain_table
      table.merge_cells(0, 0, 1, 1)

      expect do
        table.unmerge_cells(0, 1)
      end.to raise_error(Docx::Errors::InvalidMergeTarget)

      expect do
        table.unmerge_cells(1, 1)
      end.to raise_error(Docx::Errors::InvalidMergeTarget)
    end

    it 'is a no-op when unmerge_cells targets an unmerged anchor cell' do
      table = open_plain_table

      expect do
        table.unmerge_cells(1, 1)
      end.not_to raise_error

      expect(table.cell_at(1, 1).colspan).to eq(1)
      expect(table.merged?(1, 1)).to eq(false)
    end

    it 'raises InvalidMergeTarget when unmerge! is called on a non-anchor cell' do
      table = open_plain_table
      table.merge_cells(0, 0, 1, 1)

      expect do
        table.cell_at(0, 1).unmerge!
      end.to raise_error(Docx::Errors::InvalidMergeTarget)
    end

    it 'raises InvalidMergeTarget when unmerge! is called on an unmerged cell' do
      table = open_plain_table

      expect do
        table.cell_at(1, 1).unmerge!
      end.to raise_error(Docx::Errors::InvalidMergeTarget)
    end
  end
end
