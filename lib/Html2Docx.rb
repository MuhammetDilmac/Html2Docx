require 'fileutils'

require 'Html2Docx/version'
require 'Html2Docx/initialization'

module Html2Docx
  ROOT_PATH = File.expand_path(File.join(File.dirname(__FILE__), '../'))
  TEMP_PATH = Dir.mktmpdir

  def self.start(options = {})
    initialization = Initialization.new(options)
  end
end

Html2Docx.start