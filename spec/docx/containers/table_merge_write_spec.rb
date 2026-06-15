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

      temp_file = Tempfile.new(['table_merge_write', '.docx'])
      temp_path = temp_file.path
      temp_file.close

      begin
        doc.save(temp_path)

        expect do
          reopened = Docx::Document.open(temp_path)
          reopened_table = reopened.tables.first

          expect(reopened_table.row_count).to eq(3)
          expect(reopened_table.column_count).to eq(3)
          expect(reopened_table.cell_at(0, 0).colspan).to eq(2)
          expect(reopened_table.cell_at(0, 0).text).to eq('r0c0')
          expect(reopened_table.cell_at(0, 2).text).to eq('r0c2')
          expect(reopened_table.cell_at(1, 0).text).to eq('r1c0')
        end.to output('').to_stderr
      ensure
        File.delete(temp_path) if temp_path && File.exist?(temp_path)
      end
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

    it 'raises NotImplementedError for multi-row merge' do
      table = open_plain_table

      expect do
        table.merge_cells(0, 0, 1, 1)
      end.to raise_error(NotImplementedError, /not yet implemented/)
    end
  end
end
