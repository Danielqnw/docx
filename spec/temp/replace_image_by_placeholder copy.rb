#!/usr/bin/env ruby
# frozen_string_literal: true

$LOAD_PATH.unshift(File.expand_path('../../lib', __dir__))

require 'docx'
require 'fileutils'
require 'optparse'

DEFAULT_PLACEHOLDER = '{{photo_a}}'
DEFAULT_OUTPUT = 'tmp/replaced-by-placeholder.docx'
FIT_MODES = %w[cover contain stretch].freeze

def usage
  <<~TEXT
    Usage:
      ruby spec/temp/replace_image_by_placeholder.rb <docx_path> <replacement_image_path> [options]

    Options:
      --placeholder VALUE   Placeholder text (default: #{DEFAULT_PLACEHOLDER})
      --output VALUE        Output docx path (default: #{DEFAULT_OUTPUT})
      --fit VALUE           cover | contain | stretch (default: stretch)

    Example:
      ruby spec/temp/replace_image_by_placeholder.rb spec/fixtures/cover-template-cn-nostamp.docx spec/fixtures/replacement.png --placeholder "{{photo_a}}" --fit cover --output tmp/out.docx
  TEXT
end

options = {
  placeholder: DEFAULT_PLACEHOLDER,
  output: DEFAULT_OUTPUT,
  fit: 'stretch'
}

parser = OptionParser.new do |opts|
  opts.banner = usage
  opts.on('--placeholder VALUE', String) { |value| options[:placeholder] = value }
  opts.on('--output VALUE', String) { |value| options[:output] = value }
  opts.on('--fit VALUE', String) { |value| options[:fit] = value.to_s.downcase }
end

begin
  parser.parse!(ARGV)
rescue OptionParser::ParseError => e
  warn e.message
  warn usage
  exit 1
end

docx_path = ARGV[0]
replacement_path = ARGV[1]

if docx_path.nil? || docx_path.strip.empty? || replacement_path.nil? || replacement_path.strip.empty?
  warn usage
  exit 1
end

unless File.exist?(docx_path)
  warn "DOCX not found: #{docx_path}"
  exit 1
end

unless File.exist?(replacement_path)
  warn "Replacement image not found: #{replacement_path}"
  exit 1
end

unless FIT_MODES.include?(options[:fit])
  warn "Invalid --fit value: #{options[:fit]}. Allowed: #{FIT_MODES.join(', ')}"
  exit 1
end

doc = Docx::Document.open(docx_path)
placeholder = options[:placeholder]
result = doc.replace_image_by_placeholder_in_table(
  placeholder,
  replacement_path,
  fit: options[:fit].to_sym,
  cleanup_placeholder: true
)

output_path = options[:output]
output_dir = File.dirname(output_path)
FileUtils.mkdir_p(output_dir) unless output_dir == '.'
doc.save(output_path)

puts "Done: placeholder #{placeholder} -> rid #{result[:relationship_id]} (fit: #{result[:fit]})"
puts "Saved: #{output_path}"
