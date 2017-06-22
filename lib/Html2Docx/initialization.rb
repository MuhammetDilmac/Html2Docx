module Html2Docx
  class Initialization
    def initialize(options = {})
      @skell_directory = File.join(ROOT_PATH, 'skell')

      check_output_file(options[:output])
      check_html_input(options[:html])
      create_temp_directory
      check_temp_directory
      copy_skell_directory
    end

    def create_temp_directory
      @temp_path = Dir.mktmpdir
    end

    def get_temp_directory
      @temp_path
    end

    def check_temp_directory
      unless Dir.exist?(@temp_path) and File.writable?(@temp_path)
        raise "Initialization failed. Temp directory is not created success. Temp Directory: #{@temp_path}"
      end
    end

    def check_output_file(output)
      output_directory = File.dirname(output)

      unless File.writable?(output_directory)
        raise "Initialization failed. Output directory is not writable. Output Directory: #{output_directory}"
      end

      if File.exist?(output)
        raise "Initialization failed. Output file is already exist. Output File: #{output}"
      end
    end

    def check_html_input(html)
      if html.empty?
        raise 'Initialization failed. HTML must be not empty.'
      end

      begin
        Nokogiri::HTML(html)
      rescue
        raise "Initialization failed. HTML validation failed. HTML Data: #{html}"
      end
    end

    def copy_skell_directory
      FileUtils.copy_entry @skell_directory, @temp_path

      check_sync_skell_to_temp
    end

    def check_sync_skell_to_temp
      unless Dir.entries(@temp_path).length == Dir.entries(@skell_directory).length
        raise "Initialization failed. Temp directory is not syncronize to skell directory. Temp Directory: #{@temp_path}"
      end
    end
  end
end