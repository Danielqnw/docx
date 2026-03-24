#!/usr/bin/env ruby
# frozen_string_literal: true

$LOAD_PATH.unshift(File.expand_path('../../lib', __dir__))

require 'docx'
require 'fileutils'
require 'optparse'

DEFAULT_PLACEHOLDER = '{{photo_a}}'
DEFAULT_OUTPUT = 'tmp/replaced-images-by-placeholder.docx'
FIT_MODES = Docx::Document::IMAGE_FIT_MODES.map(&:to_s).freeze

def usage
  <<~TEXT
    Usage:
      ruby spec/temp/replace_images_by_placeholder_in_table.rb <docx_path> <image1,image2,...> [options]

    Options:
      --placeholder VALUE       Placeholder text (default: #{DEFAULT_PLACEHOLDER})
      --output VALUE            Output docx path (default: #{DEFAULT_OUTPUT})
      --fit VALUE               cover | contain | stretch (default: stretch)
      --max-per-row VALUE       Max images per row (default: 2)
      --width VALUE             Output width in cm (e.g. 5.0)
      --height VALUE            Output height in cm (e.g. 3.0)
      --keep-placeholder        Do not remove placeholder text

    Example:
      ruby spec/temp/replace_images_by_placeholder_in_table.rb spec/fixtures/cover-template-cn-nostamp.docx spec/fixtures/a.png,spec/fixtures/b.png,spec/fixtures/c.png --placeholder "{{photo_a}}" --fit cover --width 5.0 --height 3.0 --max-per-row 2 --output tmp/out.docx
  TEXT
end

options = {
  placeholder: DEFAULT_PLACEHOLDER,
  output: DEFAULT_OUTPUT,
  fit: 'stretch',
  max_per_row: 2,
  cleanup_placeholder: true
}

parser = OptionParser.new do |opts|
  opts.banner = usage
  opts.on('--placeholder VALUE', String) { |value| options[:placeholder] = value }
  opts.on('--output VALUE', String) { |value| options[:output] = value }
  opts.on('--fit VALUE', String) { |value| options[:fit] = value.to_s.downcase }
  opts.on('--max-per-row VALUE', Integer) { |value| options[:max_per_row] = value }
  opts.on('--width VALUE', Float) { |value| options[:width] = value }
  opts.on('--height VALUE', Float) { |value| options[:height] = value }
  opts.on('--keep-placeholder') { options[:cleanup_placeholder] = false }
end

begin
  parser.parse!(ARGV)
rescue OptionParser::ParseError => e
  warn e.message
  warn usage
  exit 1
end

docx_path = ARGV[0]
images_arg = ARGV[1]

if docx_path.nil? || docx_path.strip.empty? || images_arg.nil? || images_arg.strip.empty?
  warn usage
  exit 1
end

unless File.exist?(docx_path)
  warn "DOCX not found: #{docx_path}"
  exit 1
end

image_paths = images_arg.split(',').map(&:strip).reject(&:empty?)
if image_paths.empty?
  warn 'No image paths provided.'
  exit 1
end

missing = image_paths.reject { |path| File.exist?(path) }
unless missing.empty?
  warn "Image not found: #{missing.join(', ')}"
  exit 1
end

unless FIT_MODES.include?(options[:fit])
  warn "Invalid --fit value: #{options[:fit]}. Allowed: #{FIT_MODES.join(', ')}"
  exit 1
end

doc = Docx::Document.open(docx_path)

begin
  api_opts = {
    fit: options[:fit].to_sym,
    max_images_per_row: options[:max_per_row],
    cleanup_placeholder: options[:cleanup_placeholder]
  }
  api_opts[:width] = options[:width] if options[:width]
  api_opts[:height] = options[:height] if options[:height]

  result = doc.replace_images_by_placeholder_in_table(
    options[:placeholder],
    image_paths,
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

puts "Done: placed #{result.size} image(s) with placeholder #{options[:placeholder]}"
result.each do |item|
  puts "  row=#{item[:row_index]} slot=#{item[:slot_index]} rid=#{item[:relationship_id]} entry=#{item[:entry_path]}"
end
puts "Saved: #{output_path}"
