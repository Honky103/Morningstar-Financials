#Authored by Pang Hong Ming
#This program extracts the following Financial information about the Stock from www.morningstar.com
# 1. Revenue
# 2. Cost of Revenue (COGS)
# 3. Gross Profit
# 4. Operating Income (EBIT)
# 5. Total No. of Shares
# 6. Total Cash
# 7. Total Debt
# 8. Total Equity
# 9. Inventory
# 10. Operating Cash Flow (OCF)
# 11. Capital Expenditure (CAPEX)
# 12. Free Cash Flow (FCF)
# 
# Note1: All numerical figures are in millions (mil).
#The Financial Information will be stored in a .csv output file. 

$exc =  Read-Host -Prompt "[1]   HKEX `n[2]   NASDAQ `n[3]   NYSE `nWhich Stock Exchange do you want? Key in the number `n"

while (-not($exc -le 3 -and $exc -ge 0))
{
    $exc =  Read-Host -Prompt "Perhaps you made a mistake, please try again.`n[1]   HKEX `n[2]   NASDAQ `n[3]   NYSE `nWhich Stock Exchange do you want? Key in the number `n"
}

switch($exc) {
   1 {$exchange = 'XHKG'; break} 
   2 {$exchange = 'XNAS'; break} 
   3 {$exchange = 'XASE'; break}
#   4 {$exchange = 'XASX'; break}
}

$stc =  Read-Host -Prompt "Please enter your stock code/ticker symbol. For HKEX stocks, enter the 4 digit stock code.`n"

if ($exchange -eq 'XHKG')
{
    $stc = '0'+$stc.toString()
}

#Opens the Internet explorer and navigates to the stock's morningstar main page.
$url = "www.morningstar.com/stocks/"+ $exchange + "/" +$stc + "/quote.html"
$ie = New-Object -com internetexplorer.application;
$ie.visible = $false;
$ie.navigate($url);

Write-Host "Connecting to" $url "Please wait patiently... ..."

Start-Sleep -s 3
while($ie.busy) {Start-Sleep -s 3} #let page load

#Navigates to the page with the full Income Statement (i.e. click on all financials)
$incomeurl = $ie.document.getElementsByTagName('div') |
  Where-Object { $_.className -eq 'sal-component-footer ng-scope' } |
  ForEach-Object { $_.getElementsByTagName('a') } |
  Where-Object { $_.className -eq 'ng-binding' } |
  Select-Object -Expand href

Write-Host "Connected to" $url
Write-Host -NoNewLine "Extracting from Income Statement............."

$ie.navigate($incomeurl)

Start-Sleep -s 3
while($ie.busy) {Start-Sleep -s 3} #let page load

#Obtains the urls for the full income statement, balance sheet, and cash flow for navigation
$balanceurl = $ie.document.getElementsByTagName('ul') |
  Where-Object { $_.className -eq 'r_snav' } |
  ForEach-Object { $_.getElementsByTagName('a') } |
  Select-Object -Expand href

$tabName = "Financials"

#Create Financial Table
$table = New-Object system.Data.DataTable “$tabName”

#Gets Year values for table columns
$year = $null

$counter = 0;
$ie.document.getElementsByTagName('div') |
  Where-Object { $_.className -eq 'year column6Width109px' } |
  ForEach-Object {$year+=@{$counter=$_.textContent};$counter++}

#Define Columns of table with year values
$col1 = New-Object system.Data.DataColumn Parameter,([string])
$col2 = New-Object system.Data.DataColumn $year[0],([string])
$col3 = New-Object system.Data.DataColumn $year[1],([string])
$col4 = New-Object system.Data.DataColumn $year[2],([string])
$col5 = New-Object system.Data.DataColumn $year[3],([string])
$col6 = New-Object system.Data.DataColumn $year[4],([string])
$col7 = New-Object system.Data.DataColumn $year[5],([string])

#Add the Columns
$table.columns.add($col1)
$table.columns.add($col2)
$table.columns.add($col3)
$table.columns.add($col4)
$table.columns.add($col5)
$table.columns.add($col6)
$table.columns.add($col7)

#Extracting relevant information from the Income Statement
# 1. Revenue
# 2. Cost of Goods Sold
# 3. Gross Profit
# 4. Operating Profit
# 5. Total no. of shares

$values ='Revenue', 'Cost of Revenue','Gross Profit','Operating income (EBIT)','Weighted Shareholdings'

$elements = 'i1','i6','i10','i30','i86'
for($j=0;$j -lt 5; $j++)
{
    #Create a row in the table
    $row = $table.NewRow()
    $row.Parameter = $values[$j];

    #Extract values for Parameter
    $str = 'data_' + $elements[$j]
    $children=$ie.document.getElementById($str).childNodes
    for($i=0;$i -lt 6;$i++)
    {
        $row[$year[$i]]=$children[$i].textContent
    }
    #Add the row to the table
    $table.Rows.Add($row)
}

Write-Host "Done"
Write-Host -NoNewLine "Extracting from Balance Sheet............."

#Navigates to Balance Sheet

$ie.navigate($balanceurl[1])

Start-Sleep -s 3
while($ie.busy) {Start-Sleep -s 3} #let page load

#Extracting relevant information from the Balance Sheet
# 1. Total Cash
# 2. Inventory
# 3. CAPEX
# 4. Total Debt
# 5. Total Equity

$values = 'Total Cash','Inventory','Total Debt','Total Equity'

$elements='ttgg1','i4','ttg5','ttg8'

for($j=0;$j -lt 4; $j++)
{
    #Create a row in the table
    $row = $table.NewRow()
    $row.Parameter = $values[$j];

    #Extract values for Parameter
    $str = 'data_' + $elements[$j]
    $children=$ie.document.getElementById($str).childNodes
    for($i=0;$i -lt 5;$i++)
    {
        $row[$year[$i]]=$children[$i].textContent
    }
    #Add the row to the table
    $table.Rows.Add($row)
}

Write-Host "Done"
Write-Host -NoNewLine "Extracting from Cash Flow Statement............."

#Navigates to Cash Flow Statement

$ie.navigate($balanceurl[2])

Start-Sleep -s 3
while($ie.busy) {Start-Sleep -s 3} #let page load

#Extracting relevant information from the Cash Flow Statement
# 1. Operating Cash Flow
# 2. CAPEX
# 3. Free Cash Flow

$values = 'Operating Cash Flow','CAPEX','Free Cash Flow'

$elements='i100','i96','i97'

for($j=0;$j -lt 3; $j++)
{
    #Create a row in the table
    $row = $table.NewRow()
    $row.Parameter = $values[$j];

    #Extract values for Parameter
    $str = 'data_' + $elements[$j]
    $children=$ie.document.getElementById($str).childNodes
    for($i=0;$i -lt 6;$i++)
    {
        $row[$year[$i]]=$children[$i].textContent
    }
    #Add the row to the table
    $table.Rows.Add($row)
}

Write-Host "Done"

Write-Host -NoNewLine "Saving file................"

#Displays data in the console for fun
$table | format-table -AutoSize

#Saves the file into csv
$savefilename = "Financials_"+$stc+".csv"
$table |Export-csv $savefilename
Write-Host "Done"
$pwd = get-location
$savefilelocation = $pwd.Path + "\" + $savefilename;
Write-Host "Your file is saved at" $savefilelocation

Read-Host "Press any key to exit..."
exit
