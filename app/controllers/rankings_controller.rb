class RankingsController < ApplicationController
    require 'rubyXL'
    require 'nokogiri'
    require 'open-uri'
    require 'uri'

    def new
        if request.post?
            workbook = RubyXL::Parser.parse(params[:file].tempfile.path)
            worksheet = workbook[0]
            our_website = get_host_without_www(worksheet[1][2].value)
            google_websites = worksheet[0][3..-1].map {|cell| cell.value}.reject(&:blank?)
            
            worksheet.each_with_index do |row, row_index|
                begin
                    next if row_index == 0 # skip first row
                    if worksheet[row_index][1].present?
                        keyword = worksheet[row_index][1].value
                        google_websites.each_with_index do |website, website_index|
                            # Built proper url
                            uri = URI(website)
                            if uri.instance_of?(URI::Generic)
                                uri = URI::HTTP.build({host: uri.to_s})
                            end
                            url = uri.to_s + "/search?q=" + keyword + "&num=100" # example url - 'https://www.google.ae/search?q=2d game developer&num=100'

                            # SCRAPPPPPPEEE
                            doc = Nokogiri::HTML(open(url))
                            entries = doc.css('.g')

                            entries.each_with_index do |item, entry_index|
                                cite = get_host_without_www(item.css('cite').text)
                                if cite.include?(our_website)
                                    rank = entry_index + 1
                                    worksheet.add_cell(row_index, website_index + 3, rank.to_s)
                                    break
                                end
                                # Update rank to worksheet
                            end
                        end
                    end
                rescue Exception => e
                    puts "Row Index: " + row_index.to_s
                    puts e.message
                    break
                end
            end

            file_path = Rails.root.join('public', 'Google Keywords Ranking.xlsx')
            workbook.write(file_path)
            
            respond_to :js
        end

        def get_host_without_www(url)
            return "" unless url.present?
            encoded_url = URI.encode(url)
            uri = URI.parse(encoded_url)
            uri = URI.parse("http://#{encoded_url}") if uri.scheme.nil?
            host = uri.host.downcase
            host.start_with?('www.') ? host[4..-1] : host
        end
    end

end