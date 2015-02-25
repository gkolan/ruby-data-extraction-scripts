require 'rubygems'
require 'prawn'
require 'prawn/table'
require 'time_diff'

start_time = DateTime.now

$TOTAL_DATA		=	[]
$COLLECT_DATA 	=	[]
$STATS_DATA 	= 	[]
$STATS_DATA2	=	[]
$COUNT = 1

$PROJECT_COLOR = "428bca"
$TABLE_HEADER_TEXT_COLOR = "FFFFFF"
#"E0115F"
$TABLE_HEADER_BG = "428bca"
#"FFFAFA"
#428bca
$FONT_H1 = 30
$FONT_H2 = 18
$FONT_H3 = 14
$FONT_H4 = 12
$FONT_H5 = 10
$FONT_H6 = 8

$DATETIME = DateTime.now.strftime("%Y%m%d_%H%M")
$AUTHOR = "GautamKolan"

Prawn::Document.generate("Organization-Data_#{$AUTHOR}_#{$DATETIME}.pdf") do |pdf|
	
	#PROJECT THOUGHT PROCESS
	pdf.text "Gautam Kolan Status Report [Aug 11,2014 - Aug 15,2014]", :size => $FONT_H2, :align => :left, :color => $PROJECT_COLOR, :styles => :bold
	#, :character_spacing => 2
	pdf.stroke_horizontal_rule
	pdf.move_down 5
	$STATS_DATA = [ 
		["Contract", "[[Project Name]]", "Author", "[[Name]]"],
		["Project Manager", "[[Name]]", "Author Role", "[[Role]]"]
 	]

	$STATS_DATA2 = [ 
		["1", "ABCD", "08-09-1234", "Status Update 1"],
		["2", "XYZ", "08-09-1234", "Status Update 2"],
		["3", "PQRS", "08-09-1234", "Status Update 3"]
 	]

 	pdf.table($STATS_DATA, :column_widths => [100,170,100,170], :row_colors => ["FFFFFF","F0F8FF"], :cell_style => { size: 9 }) do |table| 
    		table.column(0).style(:align => :center, :font_style => :bold)
    		#table.column(1).style(:align => :center, :font_style => :bold)
    		table.column(2).style(:align => :center, :font_style => :bold)
	end

	$COLLECT_DATA << ["Item".force_encoding("UTF-8"), "Description".force_encoding("UTF-8"), "Start Date".force_encoding("UTF-8"), "Status".force_encoding("UTF-8")]
	
	
	pdf.move_down 20
	pdf.text "Current Work/Tasks", :size => $FONT_H3, :align => :left, :style => :bold, :color => $PROJECT_COLOR
	pdf.move_down 5
	pdf.table($COLLECT_DATA + $STATS_DATA2, :column_widths => [40,260,60,180], :header => true, :row_colors => ["FFFFFF","F0F8FF"], :cell_style => { size: 9 }) do |table| 
    		table.column(0).style(:align => :center)
    		#table.column(1).style(:align => :center, :font_style => :bold)
    		table.column(2).style(:align => :center)
    		table.row(0).style(:background_color => $TABLE_HEADER_BG, :border_width => 1,:align => :center, :font_style => :bold, :text_color => $TABLE_HEADER_TEXT_COLOR)
	end

	pdf.move_down 20
	pdf.text "Completed Work/Tasks", :size => $FONT_H3, :align => :left, :style => :bold, :color => $PROJECT_COLOR
	pdf.move_down 5
	pdf.table($COLLECT_DATA + $STATS_DATA2, :column_widths => [40,260,60,180], :header => true, :row_colors => ["FFFFFF","F0F8FF"], :cell_style => { size: 9 }) do |table| 
    		table.column(0).style(:align => :center)
    		#table.column(1).style(:align => :center, :font_style => :bold)
    		table.column(2).style(:align => :center)
    		table.row(0).style(:background_color => $TABLE_HEADER_BG, :border_width => 1,:align => :center, :font_style => :bold, :text_color => $TABLE_HEADER_TEXT_COLOR)
	end

	pdf.move_down 20
	pdf.text "Upcoming Work/Tasks", :size => $FONT_H3, :align => :left, :style => :bold, :color => $PROJECT_COLOR
	pdf.move_down 5
	pdf.table($COLLECT_DATA + $STATS_DATA2, :column_widths => [40,260,60,180], :header => true, :row_colors => ["FFFFFF","F0F8FF"], :cell_style => { size: 9 }) do |table| 
    		table.column(0).style(:align => :center)
    		#table.column(1).style(:align => :center, :font_style => :bold)
    		table.column(2).style(:align => :center)
    		table.row(0).style(:background_color => $TABLE_HEADER_BG, :border_width => 1,:align => :center, :font_style => :bold, :text_color => $TABLE_HEADER_TEXT_COLOR)
	end

	pdf.move_down 20
	pdf.text "Issues/Risks ", :size => $FONT_H3, :align => :left, :style => :bold, :color => $PROJECT_COLOR
	pdf.move_down 5
	pdf.table($COLLECT_DATA + $STATS_DATA2, :column_widths => [40,260,60,180], :header => true, :row_colors => ["FFFFFF","F0F8FF"], :cell_style => { size: 9 }) do |table| 
    		table.column(0).style(:align => :center)
    		#table.column(1).style(:align => :center, :font_style => :bold)
    		table.column(2).style(:align => :center)
    		table.row(0).style(:background_color => $TABLE_HEADER_BG, :border_width => 1,:align => :center, :font_style => :bold, :text_color => $TABLE_HEADER_TEXT_COLOR)
	end

	pdf.move_down 20
	pdf.text "Time Off ", :size => $FONT_H3, :align => :left, :style => :bold, :color => $PROJECT_COLOR
	pdf.move_down 5
	pdf.table($COLLECT_DATA + $STATS_DATA2, :column_widths => [40,260,60,180], :header => true, :row_colors => ["FFFFFF","F0F8FF"], :cell_style => { size: 9 }) do |table| 
    		table.column(0).style(:align => :center)
    		#table.column(1).style(:align => :center, :font_style => :bold)
    		table.column(2).style(:align => :center)
    		table.row(0).style(:background_color => $TABLE_HEADER_BG, :border_width => 1,:align => :center, :font_style => :bold, :text_color => $TABLE_HEADER_TEXT_COLOR)
	end
	
end