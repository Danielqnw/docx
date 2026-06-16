# frozen_string_literal: true

require 'spec_helper'
require 'docx/document'

describe 'table merge read API' do
  before(:all) do
    TableFixtureBuilder.write_all_fixtures(File.join('spec/fixtures/tables'))
  end

  def open_table(fixture_name)
    doc = Docx::Document.open(File.join('spec/fixtures/tables', fixture_name))
    doc.tables.first
  end

  def expect_plain_cell(cell)
    expect(cell.colspan).to eq(1)
    expect(cell.rowspan).to eq(1)
    expect(cell.merged?).to eq(false)
    expect(cell.merge_anchor?).to eq(false)
    expect(cell.merge_continuation?).to eq(false)
  end

  describe 'plain_3x3' do
    let(:table) { open_table('plain_3x3.docx') }

    it 'reports no merge on every logical cell' do
      (0...3).each do |row|
        (0...3).each do |col|
          expect_plain_cell(table.cell_at(row, col))
          expect(table.merged?(row, col)).to eq(false)
        end
      end
    end
  end

  describe 'horizontal_merge' do
    let(:table) { open_table('horizontal_merge.docx') }

    it 'reads merge attributes on the merged anchor and independent cells' do
      anchor = table.cell_at(0, 0)
      expect(anchor.colspan).to eq(2)
      expect(anchor.rowspan).to eq(1)
      expect(anchor.merged?).to eq(true)
      expect(anchor.merge_anchor?).to eq(true)
      expect(anchor.merge_continuation?).to eq(false)

      independent = table.cell_at(0, 2)
      expect_plain_cell(independent)

      expect(table.merged?(0, 0)).to eq(true)
      expect(table.merged?(0, 1)).to eq(true)
      expect(table.merged?(0, 2)).to eq(false)
      expect(table.merged?(1, 0)).to eq(false)
    end
  end

  describe 'vertical_merge' do
    let(:table) { open_table('vertical_merge.docx') }

    it 'reads merge attributes on the anchor and physical continuation cell' do
      anchor = table.cell_at(0, 0)
      expect(anchor.rowspan).to eq(3)
      expect(anchor.colspan).to eq(1)
      expect(anchor.merge_anchor?).to eq(true)
      expect(anchor.merged?).to eq(true)
      expect(anchor.merge_continuation?).to eq(false)

      continuation = table.rows[1].cells[0]
      expect(continuation.merge_continuation?).to eq(true)
      expect(continuation.merged?).to eq(true)
      expect(continuation.merge_anchor?).to eq(false)

      expect(table.merged?(1, 0)).to eq(true)
      expect(table.merged?(2, 0)).to eq(true)
    end
  end

  describe 'rect_merge' do
    let(:table) { open_table('rect_merge.docx') }

    it 'reads rectangle merge attributes and merged? coverage' do
      anchor = table.cell_at(0, 0)
      expect(anchor.colspan).to eq(2)
      expect(anchor.rowspan).to eq(2)
      expect(anchor.merge_anchor?).to eq(true)
      expect(anchor.merged?).to eq(true)

      expect(table.merged?(0, 0)).to eq(true)
      expect(table.merged?(0, 1)).to eq(true)
      expect(table.merged?(1, 0)).to eq(true)
      expect(table.merged?(1, 1)).to eq(true)
      expect(table.merged?(2, 2)).to eq(false)
    end
  end
end
