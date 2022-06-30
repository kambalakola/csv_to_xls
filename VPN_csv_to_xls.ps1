
$CSV_File_To_Import = "D:\work\_task\help\RptExp_acsadmin_RADIUS_Accounting_2022-06-29_15-01-00.000000072 - Copy - Copy.csv"
$CSV_File_To_Export = "D:\work\_task\help\export_VPN_sessons_trimmed_$(Get-Date -UFormat "%Y-%m-%d_%H-%m-%S").csv"
        $CSV_Data_To_Import = Import-Csv $CSV_File_To_Import -Delimiter ","
        Clear-Content $CSV_File_To_Export
        Add-Content $CSV_File_To_Export "ID,ACS_Timestamp,ACSView_Timestamp,ACS_Session_ID,User_Name,Calling_Station_ID,Acct_Session_Id,Acct_Status_Type,Acct_Session_Time,Time_In_Hours" -Encoding utf8  -notypeinformation
        
######################
foreach ($Item_From_CSV in $CSV_Data_To_Import) {

    $ID_input_array = $Item_From_CSV.ID
    
    $ACS_Timestamp_array = $Item_From_CSV.ACS_Timestamp 
	$ACS_Timestamp_array = $ACS_Timestamp_array.Substring(0, $ACS_Timestamp_array.Length -4)
    
    $ACSView_Timestamp_array = $Item_From_CSV.ACSView_Timestamp 
	$ACSView_Timestamp_array = $ACSView_Timestamp_array.Substring( 0, $ACSView_Timestamp_array.Length -4)
    
    $ACS_Session_ID_array = $Item_From_CSV.ACS_Session_ID
    $User_Name_array = $Item_From_CSV.User_Name
    $Calling_Station_ID_array = $Item_From_CSV.Calling_Station_ID
    $Acct_Session_Id_array = $Item_From_CSV.Acct_Session_Id
    $Acct_Status_Type_array = $Item_From_CSV.Acct_Status_Type
    $Acct_Session_Time_array =$Item_From_CSV.Acct_Session_Time
    


    if ($Acct_Session_Time_array -eq "null"){
		
	continue}
    $Time_In_Hours=$Acct_Session_Time_array / 86400



            $Output =New-Object -TypeName PSObject -Property @{           
            'ID' =  $ID_input_array
            'ACS_Timestamp' = $ACS_Timestamp_array
            'ACSView_Timestamp' = $ACSView_Timestamp_array
            'ACS_Session_ID' = $ACS_Session_ID_array
            'User_Name' = $User_Name_array
            'Calling_Station_ID' =$Calling_Station_ID_array
            'Acct_Session_Id' = $Acct_Session_Id_array 
            'Acct_Status_Type' = $Acct_Status_Type_array 
            'Acct_Session_Time' = $Acct_Session_Time_array
            'Time_In_Hours' = $Time_In_Hours

    } | Select-Object ID,ACS_Timestamp,ACSView_Timestamp,ACS_Session_ID,User_Name,Calling_Station_ID,Acct_Session_Id,Acct_Status_Type,Acct_Session_Time,Time_In_Hours #| Export-Csv $CSV_File_To_Export -Encoding utf8 -Append -NoTypeInformation -Delimiter ","
   

   $Output | Export-Csv $CSV_File_To_Export -Encoding utf8 -Append -NoTypeInformation -Delimiter ","
    


}

#Define locations and delimiter
$csv = $CSV_File_To_Export #Location of the source file
$xlsx = "D:\work\_task\help\RptExp_acsadmin_$(Get-Date -UFormat "%Y-%m-%d_%H-%m-%S").xlsx" #Desired location of output
$delimiter = "," #Specify the delimiter used in the file

# Create a new Excel workbook with one empty sheet
$excel = New-Object -ComObject excel.application 
$workbook = $excel.Workbooks.Add(1)
$worksheet = $workbook.worksheets.Item(1)

# Build the QueryTables.Add command and reformat the data
$TxtConnector = ("TEXT;" + $csv)
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
$query = $worksheet.QueryTables.item($Connector.name)
$query.TextFileOtherDelimiter = $delimiter
$query.TextFileParseType  = 1
$query.TextFileColumnDataTypes = ,1 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1

# Execute & delete the import query
$query.Refresh()
$query.Delete()

# Save & close the Workbook as XLSX.
$Workbook.SaveAs($xlsx,51)
$excel.Quit()