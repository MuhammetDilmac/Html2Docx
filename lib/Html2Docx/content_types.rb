module Html2Docx
  class ContentTypes
    def initialize(options = {})
      @content_type_file = File.join(options.fetch(:temp), '[Content_Types].xml')
      @content_type = File.open(@content_type_file) {|f| Nokogiri::XML(f)}
      @parts = {default: [], override: []}
      initial_parts
    end

    def initial_parts
      @content_type.root.children.each do |child|
        if child.name == 'Default'
          @parts[:default].push({extension: child.attr('Extension'), content_type: child.attr('ContentType')})
        elsif child.name == 'Override'
          @parts[:override].push({part_name: child.attr('PartName'), content_type: child.attr('ContentType')})
        end

        child.remove
      end
    end

    def add_parts(object)
      if object.fetch(:type) == 'Default'
        @parts[:default].push({extension: object.fetch(:extension), content_type: object.fetch(:content_type)})
      elsif object.fetch(:type) == 'Override'
        @parts[:override].push({part_name: object.fetch(:part_name), content_type: object.fetch(:content_type)})
      end
    end

    def render
      @parts.fetch(:default).each { |child| add_default child}
      @parts.fetch(:override).each { |child| add_override child}

      File.open(@content_type_file, 'w') {|f| f.write Helpers::NokogiriHelper.to_xml(@content_type)}
    end

    def add_default(child)
      node = Nokogiri::XML::Node.new('Default', @content_type)
      node['Extension'] = child.fetch(:extension, '')
      node['ContentType'] = child.fetch(:content_type, '')
      @content_type.root << node
    end

    def add_override(child)
      node = Nokogiri::XML::Node.new('Override', @content_type)
      node['PartName'] = child.fetch(:part_name, '')
      node['ContentType'] = child.fetch(:content_type, '')
      @content_type.root << node
    end
  end
end