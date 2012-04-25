# -*- coding: UTF-8 -*-
#require "extracttool/version"

require 'rubygems'
require 'rubyXL'
require 'iconv'
require 'optparse'
module Extracttool
  #def gem_available?(name)
  #  Gem::Specification.find_by_name(name)
  #rescue Gem::LoadError
  #  false
  #rescue
  #  Gem.available?(name)
  #end
  #
  #['nokogiri', 'zip', 'rubyXL'].each do |gem_name|
  #  unless gem_available?(gem_name)
  #    puts "Missing gem : run 'gem install #{gem_name}'"
  #    exit
  #  end
  #end


  options = {}
  OptionParser.new do |opts|
    opts.banner = "Usage: ruby extractionTool.rb [options] input_strings_file input_xlsm_file"
    options[:output_directory] = './output'
    opts.on("-o", "--output OUTPUT_DIRECTORY", "Specify the output directory") do |output_dir|
      if File.exist?(output_dir)
        if File.directory?(output_dir)
          options[:output_directory] = output_dir
        else
          puts "#{output_dir} is not a directory."
          exit
        end
      else
        puts "#{output_dir} does not exist."
        exit
      end
    end

    options[:unsecable_char] = false
    opts.on("-u", "--no_unsecable_char", "In any translation, replace whitespaces and underscores with unsecable spaces.") do |u|
      options[:unsecable_char] = true
    end

    # Just display the help
    opts.on('-h', '--help', 'Help!') do
      puts opts
      exit
    end
  end.parse!

# ----------------------- Settings -------------------------


  ADD_UNSECABLE_CHAR = options[:unsecable_char]

# This return a String that has unsecable characters instead of whitespaces and underscores,
# but only when the original string has at least one lower case AND ADD_UNSEC_CAR is true.
  def self.transform_with_unsecable_char(string)
    if (ADD_UNSECABLE_CHAR && string =~ /(?=.*[a-z])/)
      string.gsub(/[ _]+/,' ')
    else
      string
    end
  end

  path_for_inputs = '/Users/chloe/Code/SecondScreen/extracttool/input/'
  path_for_outputs = '/Users/chloe/Code/SecondScreen/extracttool/output/'

  if ARGV.count < 2
    puts "You need to specify 2 input files"
    exit
  end

  in_strings_file_path  = ARGV[0]
  in_xlsm_file_path     = ARGV[1]

  COL_B = 1
  COL_H = 7
  COL_I = 8
  COL_J = 9
  COL_P = 15
  COL_Q = 16
  COL_R = 17

  out_strings_path = {
      :H_en => "#{options[:output_directory]}/en_EU",
      :I_nl => "#{options[:output_directory]}/nl_NL",
      :P_de => "#{options[:output_directory]}/de_CH",
      :Q_fr => "#{options[:output_directory]}/fr_CH",
      :R_it => "#{options[:output_directory]}/it_CH",
  }

# All key/values hash depending on the language
  conversion_hash = {
      :H_en => {},
      :I_nl => {},
      :P_de => {},
      :Q_fr => {},
      :R_it => {},
  }

  @keys_to_translate = []

# Here, we build the new localization files with the keys we asked from the input .strings file.
  File.open(in_strings_file_path, "r:UTF-16LE:UTF-8") do |file|
    pattern = "^\"(DIC_[a-zA-Z0-9_]*)\""
    file.each do |line|
      if line =~ /#{pattern}/
        @keys_to_translate << $1
      end
    end
  end

# We get here the needed array extracted from the input xlsm file
  in_xlsm_file = RubyXL::Parser.parse(in_xlsm_file_path)
  all_strings_sheet = in_xlsm_file[3] # Because the ALL STRINGS sheet is the fourth of the file!
  @pattern_hash_conversion = "DIC_"

  all_strings_sheet.each do |row|
    # We try to get a key from the J column
    key = row[COL_J].nil? ? nil : row[COL_J].value # The key : DIC_SOMETHING
    if key.nil? # This is, no key in the J column : we try the B column
      key = row[COL_B].nil? ? nil : row[COL_B].value
    end

    if key =~ /#{@pattern_hash_conversion}/
      # Here we build the conversion_hash from the xlsm file we got
      conversion_hash[:H_en][key] = row[COL_H].nil? ? nil : row[COL_H].value
      conversion_hash[:I_nl][key] = row[COL_I].nil? ? nil : row[COL_I].value
      conversion_hash[:P_de][key] = row[COL_P].nil? ? nil : row[COL_P].value
      conversion_hash[:Q_fr][key] = row[COL_Q].nil? ? nil : row[COL_Q].value
      conversion_hash[:R_it][key] = row[COL_R].nil? ? nil : row[COL_R].value
    end

  end

  Dir::mkdir(options[:output_directory]) unless File.directory?(options[:output_directory])
# We create new Localisable.strings files with corresponding translations for asked keys
  out_strings_path.each do |col, val|
    # If the directory does not exist, we create it.
    Dir::mkdir(val) unless File.directory?(val)

    file_name = "#{val}/Localizable.strings"
    File.delete(file_name) if File.exists?(file_name)
    File.open(file_name, 'w:UTF-16LE') do |file|
      @keys_to_translate.each do |key|
        val = conversion_hash[col][key]
        # If val is nil, we try the english translation by default
        if val.nil? && col != :H_en
          val = conversion_hash[:H_en][key]
        end
        file.puts "\"#{key}\" = \"#{Extracttool.transform_with_unsecable_char(val)}\";"
      end
    end
  end

end
