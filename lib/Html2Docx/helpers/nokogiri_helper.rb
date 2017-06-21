module Html2Docx
  module Helpers
    class NokogiriHelper
      def self.to_xml(xml)
        xml.to_xml(:save_with => Nokogiri::XML::Node::SaveOptions::AS_XML | Nokogiri::XML::Node::SaveOptions::NO_DECLARATION).strip
      end
    end
  end
end
class NokogiriHelper
end