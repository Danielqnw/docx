# frozen_string_literal: true

require 'spec_helper'
require 'docx'
require 'tempfile'

describe Docx::Document, '#save round-trip' do
  fixtures_path = 'spec/fixtures'

  def document_snapshot(doc)
    {
      paragraphs: doc.paragraphs.map(&:text),
      tables: doc.tables.map do |table|
        {
          row_count: table.row_count,
          column_count: table.column_count,
          rows: table.rows.map { |row| row.cells.map(&:text) }
        }
      end
    }
  end

  def round_trip(path)
    original = Docx::Document.open(path)
    snapshot = document_snapshot(original)

    temp_file = Tempfile.new(['docx_roundtrip', '.docx'])
    temp_path = temp_file.path
    temp_file.close

    original.save(temp_path)
    reopened = Docx::Document.open(temp_path)

    [snapshot, document_snapshot(reopened), temp_path]
  ensure
    File.delete(temp_path) if temp_path && File.exist?(temp_path)
  end

  context 'with an existing paragraph fixture' do
    it 'preserves paragraph text after save and reopen' do
      before_snapshot, after_snapshot, = round_trip(File.join(fixtures_path, 'basic.docx'))

      expect(after_snapshot[:paragraphs]).to eq(before_snapshot[:paragraphs])
      expect(after_snapshot[:paragraphs]).to eq(['hello', 'world'])
    end
  end

  context 'with an existing table fixture' do
    it 'preserves table structure and cell text after save and reopen' do
      before_snapshot, after_snapshot, = round_trip(File.join(fixtures_path, 'tables.docx'))

      expect(after_snapshot[:tables].size).to eq(before_snapshot[:tables].size)
      expect(after_snapshot[:tables]).to eq(before_snapshot[:tables])
    end
  end

  context 'with a script-generated table fixture' do
    let(:generated_path) do
      path = Tempfile.new(['table_fixture_builder', '.docx']).path
      TableFixtureBuilder.plain_3x3(path)
      path
    end

    after do
      File.delete(generated_path) if generated_path && File.exist?(generated_path)
    end

    it 'opens cleanly, reads tables, and round-trips without warnings' do
      doc = nil
      expect do
        doc = Docx::Document.open(generated_path)
      end.to output('').to_stderr

      expect(doc.tables.size).to eq(1)
      table = doc.tables.first
      expect(table.row_count).to eq(3)
      expect(table.column_count).to eq(3)
      expect(table.rows.map { |row| row.cells.map(&:text) }).to eq([
        %w[r0c0 r0c1 r0c2],
        %w[r1c0 r1c1 r1c2],
        %w[r2c0 r2c1 r2c2]
      ])

      before_snapshot, after_snapshot, = round_trip(generated_path)

      expect(after_snapshot[:tables]).to eq(before_snapshot[:tables])
    end
  end
end
