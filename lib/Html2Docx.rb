require 'fileutils'
require 'nokogiri'

require 'Html2Docx/version'
require 'Html2Docx/initialization'
require 'Html2Docx/content_types'

module Html2Docx
  ROOT_PATH = File.expand_path(File.join(File.dirname(__FILE__), '../'))
  TEMP_PATH = Dir.mktmpdir

  def self.start(options = {})
    # Initialization
    initialization = Initialization.new(options)
    content_types  = ContentTypes.new(options)

    # Render
    content_types.render
  end

  def self.temp_path
    TEMP_PATH
  end

  def self.root_path
    ROOT_PATH
  end
end

Html2Docx.start