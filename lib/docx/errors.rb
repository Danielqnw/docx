module Docx
  module Errors
    StyleNotFound = Class.new(StandardError)
    StyleInvalidPropertyValue = Class.new(StandardError)
    StyleRequiredPropertyValue = Class.new(StandardError)
    InvalidMergeRange = Class.new(ArgumentError)
    MergeConflict = Class.new(StandardError)
    InvalidMergeTarget = Class.new(ArgumentError)
    ImageNotFound = Class.new(StandardError)
    ImagePlaceholderNotFound = Class.new(StandardError)
  end
end