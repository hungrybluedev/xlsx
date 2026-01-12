import xlsx { Location }
import os

fn test_large() ! {
	path := os.join_path(os.dir(@FILE), 'Free_Test_Data_1MB_XLSX.xlsx')

	document := xlsx.Document.from_file(path)!

	sheet := document.sheets[1]
	assert sheet.rows.len == 37969

	part_data := xlsx.DataFrame{
		raw_data: [
			['141', 'Felisaas', 'Female', '62', '21/05/2026', 'France'],
			['142', 'Demetas', 'Female', '63', '15/10/2028', 'France'],
			['143', 'Jeromyw', 'Female', '64', '16/08/2027', 'United States'],
			['144', 'Rashid', 'Female', '65', '21/05/2026', 'United States'],
			['145', 'Dett', 'Male', '18', '21/05/2015', 'Great Britain'],
			['146', 'Nern', 'Female', '19', '15/10/2017', 'France'],
			['147', 'Kallsie', 'Male', '20', '16/08/2016', 'France'],
			['148', 'Siuau', 'Female', '21', '21/05/2015', 'Great Britain'],
			['149', 'Shennice', 'Male', '22', '21/05/2016', 'France'],
			['150', 'Chasse', 'Female', '23', '15/10/2018', 'France'],
			['151', 'Tommye', 'Male', '24', '16/08/2017', 'United States'],
			['152', 'Dorcast', 'Female', '25', '21/05/2016', 'United States'],
			['153', 'Angelee', 'Male', '26', '21/05/2017', 'Great Britain'],
			['154', 'Willoom', 'Female', '27', '15/10/2019', 'France'],
			['155', 'Waeston', 'Male', '28', '16/08/2018', 'Great Britain'],
			['156', 'Rosma', 'Female', '29', '21/05/2017', 'France'],
			['157', 'Felisaas', 'Male', '30', '21/05/2018', 'France'],
			['158', 'Demetas', 'Female', '31', '15/10/2020', 'Great Britain'],
			['159', 'Jeromyw', 'Female', '32', '16/08/2019', 'France'],
		]
	}
	extracted_data := sheet.get_data(Location.from_encoding('A142')!, Location.from_encoding('F160')!)!
	assert part_data == extracted_data
}
