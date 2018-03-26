class RankingsController < ApplicationController
    require 'rubyXL'
    require 'open-uri'
    require 'uri'
    require 'google_search_results'

    def new
        if request.post?
            workbook = RubyXL::Parser.parse(params[:file].tempfile.path)
            worksheet = workbook[0]
            our_website = get_host_without_www(worksheet[1][2].value)
            google_websites = worksheet[0][3..-1].map {|cell| cell.value}.reject(&:blank?)
            
            worksheet.each_with_index do |row, row_index|
                next if row_index == 0 # skip first row
                if worksheet[row_index][1].present?
                    keyword = worksheet[row_index][1].value
                    google_websites.each_with_index do |website, website_index|
                        begin
                            puts "\n\n\nRow #: #{row_index + 1} - Querying '#{website}' for '#{keyword}' "
                            query_params = {
                                q: keyword,
                                google_domain: get_host_without_www(website),
                                location: "United States", 
                                num: 100
                            }
                            query = GoogleSearchResults.new query_params
                            
                            hash_results = query.get_hash
                            puts "Got #{hash_results[:organic_results].count} results"

                            puts "Looping over each result item..."
                            hash_results[:organic_results].each_with_index do |item, item_index|
                                puts "Result item # #{item_index}"
                                link = get_host_without_www(item[:link])
                                puts "Matching result link '#{link}' with our website '#{our_website}'..."
                                if link.include?(our_website)
                                    puts "MATCHED!!!"
                                    rank_and_url_str = "(#{item[:position].to_s}) #{item[:link]}"
                                    puts "Updating rank to worksheet..."
                                    worksheet.add_cell(row_index, website_index + 3, rank_and_url_str)
                                    break
                                end
                                if hash_results[:organic_results].last == item
                                    worksheet.add_cell(row_index, website_index + 3, "No rank")
                                end
                            end
                        rescue Exception => e
                            puts e.message
                            worksheet.add_cell(row_index, website_index + 3, e.message)
                        end
                    end
                end
            end

            file_path = Rails.root.join('public', 'Google Keywords Ranking.xlsx')
            workbook.write(file_path)
            
            respond_to :js
        end

        def get_host_without_www(url)
            # This will prepend http to a url. For example: google.com will become http://google.com
            # uri = URI(website)
            # if uri.instance_of?(URI::Generic)
            #     uri = URI::HTTP.build({host: uri.to_s})
            # end
            # uri.to_s
            return "" unless url.present?
            encoded_url = URI.encode(url)
            uri = URI.parse(encoded_url)
            uri = URI.parse("http://#{encoded_url}") if uri.scheme.nil?
            host = uri.host.downcase
            host.start_with?('www.') ? host[4..-1] : host
        end
    end

end