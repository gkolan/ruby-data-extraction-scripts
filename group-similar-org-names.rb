require 'rubygems'
require 'time'
require 'csv'
require 'mail'
require 'spreadsheet'
require 'roo'
require 'amatch'
include Amatch

$MASTER_ARRAY 		= []
$COPY_ARRAY 		= []
$COMPLETE_FILTER 	= []
count = 0
$THRESHOLD = 0.91
$TIMENOW = Time.now.strftime("%Y-%m-%d_at_%H-%M")
=begin
spreadsheet_master = Roo::Excelx.new("master.xlsx") 
spreadsheet_master.default_sheet = spreadsheet_master.sheets.first
(2..spreadsheet_master.last_row).each do |i|
	master_id 		=  spreadsheet_master.cell('A',i).to_i
	master_name 	=  spreadsheet_master.cell('B',i)
	$MASTER_ARRAY << [master_id, master_name]
end
puts $MASTER_ARRAY.size

spreadsheet_copy = Roo::Excelx.new("copy.xlsx") 
spreadsheet_copy.default_sheet = spreadsheet_copy.sheets.first
(1..spreadsheet_copy.last_row).each do |i|
	copy_id 		=  spreadsheet_copy.cell('A',i).to_i
	copy_name 		=  spreadsheet_copy.cell('B',i)
	$COPY_ARRAY << [copy_id, copy_name]
end
puts $COPY_ARRAY.size


CSV.open("new4.csv", "w") do |csv|
	
	csv << ["Group","Org ID","Org Name","Type"]
	count = count + 1
	puts count.to_s + ' : Headline'.to_s 
	
	while $MASTER_ARRAY.size > 0
		$FILTER_ARRAY = []
		first_value = $MASTER_ARRAY.shift
		ma_id = first_value[0]
		ma_name = first_value[1]
		csv << [ma_id,ma_id,ma_name,"Master"]
		count = count + 1
		puts count.to_s + ' : Master'.to_s 
		m = JaroWinkler.new("#{ma_name}")
		$COPY_ARRAY.map do |ca|
			ca_id 	= ca[0]
			ca_name = ca[1]
			value = m.match("#{ca_name}")
			if value > 0.91
				csv << [ma_id,ca_id,ca_name,"Copy"]
				count = count + 1
				puts count.to_s + ' : Copy'.to_s 
				$FILTER_ARRAY << ca
				$COMPLETE_FILTER << ca
			end
		end
		$COPY_ARRAY = $COPY_ARRAY - $FILTER_ARRAY
	end
	#puts count
	puts $COPY_ARRAY.size

	while $COPY_ARRAY.size > 0
		$FILTER_ARRAY = []
		first_value = $COPY_ARRAY.shift
		copy_master_id 		= first_value[0]
		copy_master_name 	= first_value[1]
		csv << [copy_master_id,copy_master_id,copy_master_name,"Master"]
		count = count + 1
		puts count.to_s + ' : Master'.to_s 
		ca_m = JaroWinkler.new("#{copy_master_name}")
		$COPY_ARRAY.each do |new_ca|
			new_ca_id 	= new_ca[0]
			new_ca_name = new_ca[1]
			ca_value = ca_m.match("#{new_ca_name}")
			if ca_value > 0.91
				csv << [copy_master_id,new_ca_id,new_ca_name,"Copy"]
				count = count + 1
				puts count.to_s + ' : Copy'.to_s 
				$FILTER_ARRAY << new_ca
			end
		end 
		$COPY_ARRAY = $COPY_ARRAY - $FILTER_ARRAY
	end
end
	
CSV.open("EMAIL_DOMAINS.csv", "w") do |csv|
	csv << ["EMAIL_DOMAIN", "UNIQ_ORG_COUNT", "TOTAL_ORG_COUNT", "UNIQ_ORGS", "TOTAL_ORGS"]
	$LINE_ARRAY.group_by{|i,j| i}.map { |(i,j)| [i, j.map { |(a,b)|  b} ]}.each do |line|
		csv << [line[0],line[1].uniq.size,line[1].size,line[1].uniq,line[1]]
		puts "**" + count.to_s
		count = count + 1
	end
end

=end

spreadsheet_copy = Roo::Excelx.new("EXTR0001.xlsx") 
spreadsheet_copy.default_sheet = spreadsheet_copy.sheets.first
(1..spreadsheet_copy.last_row).each do |i|
	copy_id 		=  spreadsheet_copy.cell('A',i).to_i
	copy_name 		=  spreadsheet_copy.cell('B',i)
	$COPY_ARRAY << [copy_id, copy_name]
end
puts $COPY_ARRAY.size
CSV.open("All_Data_#{$THRESHOLD}_#{$TIMENOW}.csv", "w") do |csv|
CSV.open("Only_Master_#{$THRESHOLD}_#{$TIMENOW}.csv", "w") do |csvm|
CSV.open("Only_Duplicates_#{$THRESHOLD}_#{$TIMENOW}.csv", "w") do |csvd|
	csv << ["Group","Org ID","Org Name","Type"]
	csvm << ["Group","Org ID","Org Name","Type"]
	csvd << ["Group","Org ID","Org Name","Type"]
	count = count + 1
	puts count.to_s + ' : Headline'.to_s
	while $COPY_ARRAY.size > 0
		$FILTER_ARRAY 	= []
		$COLLECTION		= []
		first_value = $COPY_ARRAY.shift
		copy_master_id 		= first_value[0]
		copy_master_name 	= first_value[1]
		csv << [copy_master_id,copy_master_id,copy_master_name,"Master"]
		$COLLECTION  << [copy_master_id,copy_master_id,copy_master_name,"Master"]
		count = count + 1
		puts count.to_s + ' : Master'.to_s 
		ca_m = JaroWinkler.new("#{copy_master_name}")
		$COPY_ARRAY.each do |new_ca|
			new_ca_id 	= new_ca[0]
			new_ca_name = new_ca[1]
			ca_value = ca_m.match("#{new_ca_name}")
			if ca_value > $THRESHOLD
				csv << [copy_master_id,new_ca_id,new_ca_name,"Copy"]
				$COLLECTION << [copy_master_id,new_ca_id,new_ca_name,"Copy"]
				count = count + 1
				puts count.to_s + ' : Copy'.to_s 
				$FILTER_ARRAY << new_ca
			end
		end
		$COLLECTION << ["","","",""]
		collection_size =  $COLLECTION.size
		#print $COLLECTION
		#puts
		puts "collection_size : " + collection_size.to_s
		if collection_size > 2
			$COLLECTION.each do |item|
				csvd << item
			end
		else
			csvm << $COLLECTION[0]
		end

		$COPY_ARRAY = $COPY_ARRAY - $FILTER_ARRAY
	end
end
end
end