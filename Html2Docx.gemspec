# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'Html2Docx/version'

Gem::Specification.new do |spec|
  spec.name          = 'Html2Docx'
  spec.version       = Html2Docx::VERSION
  spec.authors       = ['MuhammetDilmac']
  spec.email         = ['iletisim@muhammetdilmac.com.tr']

  spec.summary       = 'HTML çıktısından Docx oluşturmayı sağlayan ruby' \
                       'kütüphanesi'
  spec.description   = 'Kendisine özel olarak oluşturulan html çıktısını ' \
                        'işleyerek bu çıktıdan Docx üretmeyi sağlayan ruby ' \
                        'kütüphanesi'
  spec.homepage      = 'https://www.github.com/MuhammetDilmac/Html2Docx'
  spec.license       = 'MIT'

  spec.files         = `git ls-files -z`.split("\x0").reject do |f|
    f.match(%r{^(test|spec|features)/})
  end
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ['lib']

  spec.add_development_dependency 'bundler', '~> 1.15'
  spec.add_development_dependency 'rake', '~> 10.0'
  spec.add_development_dependency 'rspec', '~> 3.0'
  spec.add_development_dependency 'nokogiri', '~> 1.6', '>= 1.6.8'
end
