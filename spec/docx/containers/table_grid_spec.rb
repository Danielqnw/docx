# frozen_string_literal: true

require 'spec_helper'
require 'docx/document'

describe Docx::Elements::Containers::TableGrid do
  FIXTURES_DIR = File.join('spec/fixtures/tables')

  before(:all) do
    TableFixtureBuilder.write_all_fixtures(FIXTURES_DIR)
  end

  def open_table(fixture_name)
    doc = Docx::Document.open(File.join(FIXTURES_DIR, fixture_name))
    doc.tables.first
  end

  def grid_for(table)
    table.send(:grid)
  end

  shared_examples 'columns align with cell_at' do |fixture_name|
    it 'maps columns[col].cells[row] to cell_at(row, col)' do
      table = open_table(fixture_name)

      expect(table.columns.size).to eq(table.column_count)

      table.columns.each_with_index do |column, col|
        expect(column.cells.size).to eq(table.row_count)

        column.cells.each_with_index do |cell, row|
          expected = table.cell_at(row, col)
          next if expected.nil?

          expect(cell.to_s).to eq(expected.to_s)
          expect(cell.instance_variable_get(:@node)).to eq(expected.instance_variable_get(:@node))
        end
      end
    end
  end

  describe 'plain_3x3' do
    let(:table) { open_table('plain_3x3.docx') }
    let(:grid) { grid_for(table) }

    it 'returns anchor text at each logical coordinate' do
      (0...3).each do |row|
        (0...3).each do |col|
          expect(table.cell_at(row, col).text).to eq("r#{row}c#{col}")
        end
      end
    end

    it 'yields each anchor cell once via each_cell' do
      anchors = []
      table.each_cell { |cell, row, col| anchors << [cell.text, row, col] }

      expect(anchors.size).to eq(9)
      expect(anchors).to match_array(
        (0...3).flat_map { |row| (0...3).map { |col| ["r#{row}c#{col}", row, col] } }
      )
    end

    it 'returns nil for out-of-bounds coordinates' do
      expect(table.cell_at(3, 0)).to be_nil
      expect(table.cell_at(0, 3)).to be_nil
    end

    it 'has expected row and column counts' do
      expect(table.row_count).to eq(3)
      expect(table.column_count).to eq(3)
    end

    include_examples 'columns align with cell_at', 'plain_3x3.docx'
  end

  describe 'horizontal_merge' do
    let(:table) { open_table('horizontal_merge.docx') }
    let(:grid) { grid_for(table) }

    it 'maps merged columns to the same anchor cell' do
      anchor_node = table.cell_at(0, 0).instance_variable_get(:@node)
      expect(table.cell_at(0, 1).instance_variable_get(:@node)).to eq(anchor_node)
      expect(table.cell_at(0, 0).text).to eq(table.cell_at(0, 1).text)
    end

    it 'sets anchor colspan to merged width' do
      slot = grid.cell_at(0, 0)
      expect(slot.colspan).to eq(2)
    end

    it 'keeps unmerged cells independent' do
      merged_node = table.cell_at(0, 0).instance_variable_get(:@node)
      independent = table.cell_at(0, 2)

      expect(independent.instance_variable_get(:@node)).not_to eq(merged_node)
      expect(independent.text).to eq('r0c2')
    end

    it 'has expected row and column counts' do
      expect(table.row_count).to eq(3)
      expect(table.column_count).to eq(3)
    end

    include_examples 'columns align with cell_at', 'horizontal_merge.docx'
  end

  describe 'vertical_merge' do
    let(:table) { open_table('vertical_merge.docx') }
    let(:grid) { grid_for(table) }

    it 'maps continuation rows back to the top anchor' do
      anchor_node = table.cell_at(0, 0).instance_variable_get(:@node)
      expect(table.cell_at(1, 0).instance_variable_get(:@node)).to eq(anchor_node)
      expect(table.cell_at(2, 0).instance_variable_get(:@node)).to eq(anchor_node)
    end

    it 'sets anchor rowspan to merged height' do
      slot = grid.cell_at(0, 0)
      expect(slot.rowspan).to eq(3)
    end

    it 'has expected row and column counts' do
      expect(table.row_count).to eq(3)
      expect(table.column_count).to eq(3)
    end

    include_examples 'columns align with cell_at', 'vertical_merge.docx'
  end

  describe 'rect_merge' do
    let(:table) { open_table('rect_merge.docx') }
    let(:grid) { grid_for(table) }

    it 'returns the same anchor for every coordinate in the merged rectangle' do
      anchor_node = table.cell_at(0, 0).instance_variable_get(:@node)

      (0..1).each do |row|
        (0..1).each do |col|
          expect(table.cell_at(row, col).instance_variable_get(:@node)).to eq(anchor_node)
        end
      end
    end

    it 'sets anchor colspan and rowspan for the rectangle' do
      slot = grid.cell_at(0, 0)
      expect(slot.colspan).to eq(2)
      expect(slot.rowspan).to eq(2)
    end

    it 'has expected row and column counts' do
      expect(table.row_count).to eq(3)
      expect(table.column_count).to eq(3)
    end

    include_examples 'columns align with cell_at', 'rect_merge.docx'
  end
end
