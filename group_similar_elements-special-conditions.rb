require 'rubygems'
require 'roo'
require 'csv'

$READ_FILEPATH 			= 'OrgNames.xlsx'
$SHEET_NAME 			= 'Sheet1'
$SHEET_HEADER 			= true
$ELEMENTS_COLUMN_INFO	= [5] #[Column no,more info to be added later] 
$GROUP_COLUMN_INFO		= [1] #[Column no,more info to be added later]
$START_ROW_NUM			= 1
$HEADER_INFO			= ""
$SORT_ARRAY 			= []
$GROUP_ARRAY 			= []
$PRINT_ARRAY 			= []
$INTEGER_TIME			= Time.now.to_i
$READ_COUNT				= 0
$WRITE_COUNT			= 0

book =  Roo::Excelx.new($READ_FILEPATH)
book.default_sheet = $SHEET_NAME
if $SHEET_HEADER 
	$HEADER_INFO = book.row(1)
	$START_ROW_NUM = 2
end
puts " *** READ START *** "
($START_ROW_NUM..book.last_row).each do |i|
	$SORT_ARRAY << [book.cell(i,$GROUP_COLUMN_INFO[0]).to_i,book.cell(i,$ELEMENTS_COLUMN_INFO[0]).to_i]
end
puts " *** READ END *** "
puts
puts " *** PROCESS START *** "
$GROUP_ARRAY = $SORT_ARRAY.group_by{|(x,y)| x }.map{|(x,y)| y }.map{|x| [x.map {|(x,y)| x}.flatten.uniq,x.map {|(x,y)| [y]}.flatten.uniq.join(", ") ] }
puts " *** PROCESS END *** "
puts
puts " *** WRITE START *** "
CSV.open("group_similar_elements_for_#{$READ_FILEPATH}_#{$INTEGER_TIME}.csv", "wb") do |csv|
	$GROUP_ARRAY.each do |value|
  		csv << [value[0][0]," [ " + value[1] + " ] "]
  		$WRITE_COUNT += 1
  		puts $WRITE_COUNT
  	end
end
puts " *** WRITE END *** "
puts