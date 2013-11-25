# USAGE:
# ruby proteins_to_genes.rb ../data/list_DAVID.xlsx ../results/genes_list.xlsx

# retrieve genes from the protein description, keep only unique proteins

require 'rubygems'
require 'rubyXL'
require 'axlsx'

ifile = ARGV[0]
ofile = ARGV[1]

# initialize arguments
workbook = RubyXL::Parser.parse(ifile)

# read the lists (format: prot_acc	prot_desc	pep_score	pep_expect	pep_seq)
unique_list = Hash.new { |h,k| h[k] = [] }
worksheet = workbook[0]
array = worksheet.extract_data
array.each do |row|
	if !row[0].include? "prot_acc" 
		if !unique_list.has_key?(row[0])
			if row[1].include? "GN="
				genename = row[1].split("GN=")[1].split(" ")[0].to_s
				row_updated = row << genename
				unique_list[row[0]] = row_updated
			end
		else
		end
	end
end

# output
results_xlsx = Axlsx::Package.new
results_wb = results_xlsx.workbook

# create sheet
results_wb.add_worksheet(:name => "genes list") do |sheet|
	sheet.add_row ["prot_acc", "prot_desc", "prot_score", "pep_exp_mz", "pep_exp_mr", "pep_exp_z", "pep_delta", "pep_score", "pep_expect", "pep_seq", "score (combine)", "genename"]
	unique_list.each_key do |protein|
		prot_acc = unique_list[protein][0]
		prot_desc = unique_list[protein][1]
		prot_score = unique_list[protein][2]
		pep_exp_mz = unique_list[protein][3]
		pep_exp_mr = unique_list[protein][4]
		pep_exp_z = unique_list[protein][5]
		pep_delta = unique_list[protein][6]
		pep_score = unique_list[protein][7]
		pep_expect = unique_list[protein][8]
		pep_seq = unique_list[protein][9]
		score = unique_list[protein][10]
		genename = unique_list[protein][11]
		row = sheet.add_row [prot_acc, prot_desc, prot_score, pep_exp_mz, pep_exp_mr, pep_exp_z, pep_delta, pep_score, pep_expect, pep_seq, score, genename]
	end
end

# write xlsx file
results_xlsx.serialize(ofile)

