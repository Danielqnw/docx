#!/usr/bin/env ruby
# frozen_string_literal: true

$LOAD_PATH.unshift(File.expand_path('../../lib', __dir__))

require 'docx'
require 'fileutils'
require 'optparse'

DEFAULT_PLACEHOLDER = '{{photo_a}}'
DEFAULT_OUTPUT = 'tmp/replaced-by-placeholder.docx'
FIT_MODES = Docx::Document::IMAGE_FIT_MODES.map(&:to_s).freeze

def usage
  <<~TEXT
    Usage:
      ruby spec/temp/replace_image_by_placeholder.rb <docx_path> <replacement_image_path> [options]

    Options:
      --placeholder VALUE   Placeholder text (default: #{DEFAULT_PLACEHOLDER})
      --output VALUE        Output docx path (default: #{DEFAULT_OUTPUT})
      --fit VALUE           cover | contain | stretch (default: stretch)
      --width VALUE         Output width in cm (e.g. 5.0)
      --height VALUE        Output height in cm (e.g. 3.0)

    Example:
      ruby spec/temp/replace_image_by_placeholder.rb spec/fixtures/cover-template-cn-nostamp.docx spec/fixtures/replacement.png --placeholder "{{photo_a}}" --fit cover --width 5.0 --height 3.0 --output tmp/out.docx
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
  opts.on('--width VALUE', Float) { |value| options[:width] = value }
  opts.on('--height VALUE', Float) { |value| options[:height] = value }
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
api_opts = { fit: options[:fit].to_sym, cleanup_placeholder: true }
api_opts[:width] = options[:width] if options[:width]
api_opts[:height] = options[:height] if options[:height]

begin
  result = doc.replace_image_by_placeholder_in_table(
    placeholder,
    replacement_path,
    api_opts
  )
rescue Docx::Errors::ImagePlaceholderNotFound, Docx::Errors::ImageNotFound, ArgumentError => e
  warn e.message
  exit 1
end

output_path = options[:output]
output_dir = File.dirname(output_path)
FileUtils.mkdir_p(output_dir) unless output_dir == '.'
doc.save(output_path)

size_info = [options[:width] ? "w=#{options[:width]}cm" : nil, options[:height] ? "h=#{options[:height]}cm" : nil].compact.join(' ')
puts "Done: placeholder #{placeholder} -> rid #{result[:relationship_id]} (fit: #{result[:fit]}) #{size_info}"
puts "Saved: #{output_path}"
