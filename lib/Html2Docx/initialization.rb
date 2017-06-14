module Html2Docx
  class Initialization
    def initialize(options = {})
      @skell_directory = File.join(ROOT_PATH, 'skell')

      check_temp_directory
      copy_skell_directory
    end

    def check_temp_directory
      unless Dir.exist?(TEMP_PATH) and File.writable?(TEMP_PATH)
        raise "Initialization failed. Temp directory is not created success. Temp Directory: #{TEMP_PATH}"
      end
    end

    def copy_skell_directory
      FileUtils.copy_entry @skell_directory, TEMP_PATH

      check_sync_skell_to_temp
    end

    def check_sync_skell_to_temp
      unless Dir.entries(TEMP_PATH).length == Dir.entries(@skell_directory).length
        raise "Initialization failed. Temp directory is not syncronize to skell directory. Temp Directory: #{TEMP_PATH}"
      end
    end
  end
end