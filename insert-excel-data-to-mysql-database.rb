require 'rubygems'
require 'mysql2'
require 'roo'

db = Mysql2::Client.new(:host => "localhost", :username => "root",  :password => "root")


sp = Roo::Excelx.new("EXTR002A.xlsx")
sp.default_sheet = sp.sheets.first

(1..sp.last_row).each do |i|
	time_now = Time.now.strftime("%F %T")

	id_value 				= i
	adr_id_value 			= "#{sp.cell('A',i).to_i}"
	adr_sqnc_num_value 		= "#{sp.cell('B',i).to_i}"
	cntry_cd_value 			= "#{sp.cell('C',i)}"
	state_cd_value 			= "#{sp.cell('D',i)}"
	univ_org_id_value 		= "#{sp.cell('E',i).to_i}"
	adr_cnct_name_value 	= "#{sp.cell('F',i)}" 
	adr_city_name_value 	= "#{sp.cell('G',i)}"
	adr_zip_cd_value 		= "#{sp.cell('H',i).to_i}"
	adr_phn_num_value 		= "#{sp.cell('I',i).to_i}"
	adr_extnsn_num_value 	= "#{sp.cell('J',i).to_i}"
	adr_email_name_value 	= "#{sp.cell('K',i)}"
	entry_ts_value 			= "#{sp.cell('L',i)}"
	crnt_adr_sw_value 		= "#{sp.cell('M',i)}"
	user_id_value 			= 1
	organization_id_value 	= univ_org_id_value
	is_edited_value 		= 0
	is_archived_value 		= 0
	created_at_value 		= "#{time_now}"
	updated_at_value 		= "#{time_now}"

	db.query "INSERT INTO corbett_development.contacts (id,adr_id,adr_sqnc_num,cntry_cd,state_cd,univ_org_id,adr_cnct_name,adr_city_name,adr_zip_cd,adr_phn_num,adr_extnsn_num,adr_email_name,entry_ts,crnt_adr_sw,user_id,organization_id,is_edited,is_archived,created_at,updated_at) 
				VALUES ('#{id_value}','#{adr_id_value}','#{adr_sqnc_num_value}','#{cntry_cd_value}','#{state_cd_value}','#{univ_org_id_value}','#{adr_cnct_name_value}','#{adr_city_name_value}','#{adr_zip_cd_value}','#{adr_phn_num_value}','#{adr_extnsn_num_value}','#{adr_email_name_value}','#{entry_ts_value}','#{crnt_adr_sw_value}','#{user_id_value}','#{organization_id_value}','#{is_edited_value}','#{is_archived_value}','#{time_now}','#{time_now}')"
  	puts i
end

s = Roo::Excelx.new("EXTR0001.xlsx")
s.default_sheet = s.sheets.first

(1..s.last_row).each do |i|
	time_now = Time.now.strftime("%F %T")
	
	id_value 			= 	i
	univ_org_id_value 	=	"#{s.cell('A',i).to_i}"
	univ_org_name_value	=	"#{s.cell('B',i)}".gsub("'","").to_s
	user_id_value		=	1
	is_master_org_value	=	0
	is_edited_value		=	0
	is_archived_value	=	0
	created_at_value	=	"#{time_now}"
	updated_at_value	=	"#{time_now}"

	db.query "INSERT INTO corbett_development.organizations (id,univ_org_id,univ_org_name,user_id,is_master_org,is_edited,is_archived,created_at,updated_at) 
				VALUES ('#{id_value.to_i}','#{univ_org_id_value.to_i}','#{univ_org_name_value}','#{user_id_value.to_i}','#{is_master_org_value.to_i}','#{is_edited_value.to_i}','#{is_archived_value.to_i}','#{time_now}','#{time_now}')"
  	puts i
end


# 235
# 623