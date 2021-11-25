require "roo-xls"
require "pry"

class Parserxls
  def initialize(file_xls)
    @file = Roo::Excel.new file_xls
    @hash ||= {}
    @hash_att ||= {}
  end

  def parse
    @file.each_with_index do |cell, index|
      case index
      when 8
        cell.each_with_index { |item, index_line| @hash_att[index_line] = item unless item.nil? }
      when 10..cell.size
        @hash_att.each do |key, value|
          @hash[cell[0]] ||= {}
          @hash[cell[0]][value] = cell[key] if key > 0
        end
      end
    end
    @hash
  end

  def print_console
    @hash.each do |name, data|
      puts name
      @hash[name].each do |name_data, value_data|
        print "#{name_data}: #{value_data}"
        puts
      end
      puts "================================================================================================================="
    end
  end
end

data = Parserxls.new(ARGV[0])
data.parse
data.print_console
