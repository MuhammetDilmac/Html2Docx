module Html2Docx
  class ContentTypes
    def initialize(options = {})
      @content_type_file = File.join(TEMP_PATH, '[Content_Types].xml')
      @content_type = File.open(@content_type_file) { |f| Nokogiri::XML(f) }
      @types = @content_type.css('Types')
      @parts = { default: [], override: [] }

      initial_parts
    end

    def initial_parts
      @types.children.each do |child|
        if child.name == 'Default'
          @parts[:default].push( { extension: child.attr('Extension'), content_type: child.attr('ContentType') } )
        elsif child.name == 'Override'
          @parts[:override].push( { part_name: child.attr('PartName'), content_type: child.attr('ContentType') } )
        end
      end

      @types.remove
    end

    def add_parts(object)
      if object.fetch(:type) == 'Default'
        @parts[:default].push( { extension: object.fetch(:extension), content_type: object.fetch(:content_type) } )
      elsif object.fetch(:type) == 'Override'
        @parts[:override].push( { part_name: object.fetch(:part_name), content_type: object.fetch(:content_type) } )
      end
    end

    def render
      @parts.fetch(:default).each { |child| add_child child }
      @parts.fetch(:override).each { |child| add_child child }
    end

    def add_child(child)
      node = Nokogiri::XML::Node.new('Default', @content_type)
      node['Extension'] = child.fetch(:extension)
      node['ContentType'] = child.fetch(:content_type)

      @types.add_child(node)
    end
  end
end