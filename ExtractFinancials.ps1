#Authored by Pang Hong Ming
###### This program extracts the following Financial information about the Stock from www.morningstar.com ######
# 1. Revenue
# 2. Cost of Revenue (COGS)
# 3. Gross Profit
# 4. Operating Income (EBIT)
# 5. Total No. of Shares
# 6. Account Receivables
# 7. Account Payables
# 8. Total Cash
# 9. Total Debt
# 10. Total Equity
# 11. Inventory
# 12. Operating Cash Flow (OCF)
# 13. Capital Expenditure (CAPEX)
# 14. Free Cash Flow (FCF)
# 
# Note1: All numerical figures are in millions (mil).
#The Financial Information will be stored in a .csv output file. 

$ie = New-Object -com internetexplorer.application;
$ie.visible = $false;
$tabName = 'Financials'
$financial_table = New-Object system.Data.DataTable â€œ$tabNameâ€

$gim_score = 0;
$year = $null
$incomeurl = $null
$Timeout = 60 #set timeout of web request to 60s

#Requests user to select stock exchange and input stock code
Do{$exc =  Read-Host -Prompt "[1]   HKEX `n[2]   NASDAQ `n[3]   NYSE AMERICAN `n[4]   NYSE (Currently not working for banks and financial institutions) `nWhich Stock Exchange do you want? Key in the number `n"} while ((1..4) -notcontains $exc)

switch($exc) 
{
   1 {$exchange = 'XHKG'; break} 
   2 {$exchange = 'XNAS'; break} 
   3 {$exchange = 'XASE'; break}
   4 {$exchange = 'XNYS'; break}
}

$stc =  Read-Host -Prompt "Please enter your stock code/ticker symbol. For HKEX stocks, enter the 4 digit stock code.`n"

if ($exchange -eq 'XHKG')
{
    $stc = '0'+$stc.toString()
}

#Opens the Internet explorer and navigates to the stock's morningstar page.
$url = "www.morningstar.com/stocks/"+ $exchange + "/" +$stc + "/quote.html"
$ie.navigate($url);

#let page load
Write-Host "Connecting to" $url "please wait patiently... ..."
Start-Sleep -s 3
while($ie.busy) {Start-Sleep -s 3}
Start-Sleep -s 3

$timer = [Diagnostics.Stopwatch]::StartNew()

#Navigates to the page with the full Income Statement (i.e. click on all financials)
while(($timer.Elapsed.TotalSeconds -le $Timeout) -and $incomeurl -eq $null)
{
    $incomeurl = $ie.document.getElementsByTagName('div') |
    Where-Object { $_.className -eq 'sal-component-footer ng-scope' } |
    ForEach-Object { $_.getElementsByTagName('a') } |
    Where-Object { $_.className -eq 'ng-binding' } |
    Select-Object -Expand href
}

#Checks if the request timeout, if request is timeout, outputs that the stock code is invalid. (highly likely wrong stockcode)
$timer.Stop()
if ($timer.Elapsed.TotalSeconds -gt $Timeout) {
     Write-Host 'Could not connect to' $url ', please check your stock code/ticker symbol' -ForegroundColor Red
     Read-Host "Press Enter to exit"
     exit
 }

Write-Host "Connected to" $url
Write-Host -NoNewLine "Extracting from Income Statement............."

$ie.navigate($incomeurl)

#let page load
Start-Sleep -s 3
while($ie.busy) {Start-Sleep -s 3}
Start-Sleep -s 3

#Obtains the urls for the full income statement, balance sheet, and cash flow for navigation
$balanceurl = $ie.document.getElementsByTagName('ul') |
  Where-Object { $_.className -eq 'r_snav' } |
  ForEach-Object { $_.getElementsByTagName('a') } |
  Select-Object -Expand href

#Gets the last 5 year values from Morningstar's table
while($year -eq $null)
{
    $counter = 0;
    $ie.document.getElementsByTagName('div') |
    Where-Object { $_.className -eq 'year column6Width109px' } |
    ForEach-Object {$year+=@{$counter=$_.textContent};$counter++}
}

#Define Columns of table with year values
$col1 = New-Object system.Data.DataColumn Millions,([string])
$col2 = New-Object system.Data.DataColumn $year[0],([string])
$col3 = New-Object system.Data.DataColumn $year[1],([string])
$col4 = New-Object system.Data.DataColumn $year[2],([string])
$col5 = New-Object system.Data.DataColumn $year[3],([string])
$col6 = New-Object system.Data.DataColumn $year[4],([string])
$col7 = New-Object system.Data.DataColumn $year[5],([string])

$financial_table.columns.add($col1)
$financial_table.columns.add($col2)
$financial_table.columns.add($col3)
$financial_table.columns.add($col4)
$financial_table.columns.add($col5)
$financial_table.columns.add($col6)
$financial_table.columns.add($col7)

#Extracting relevant information from the Income Statement
# 1. Revenue
# 2. Cost of Goods Sold
# 3. Gross Profit
# 4. Operating Profit
# 5. Total no. of shares

$values ='Revenue', 'Cost of Revenue','Gross Profit','Operating income (EBIT)','Weighted Shareholdings'

#Element ID of the information from the HTML code
$elements = 'i1','i6','i10','i30','i86'
for($j=0;$j -lt 5; $j++)
{
    #Create a row in the table
    $row = $financial_table.NewRow()
    $row.Millions = $values[$j];

    #Extract values for Parameter
    $str = 'data_' + $elements[$j]
    $children=$ie.document.getElementById($str).childNodes
    for($i=0;$i -lt 6;$i++)
    {
        $temp = $null;
        $temp=$children[$i].textContent
        if ($temp.startsWith('('))
        {
          $temp = $temp -replace '[()]',"";
          $temp = 0-[double]$temp;  
        } 
        if ($temp -eq '—')
        {
            $temp = 0;
        }
        $row[$year[$i]] = $temp;
    }
    #Add the row to the table
    $financial_table.Rows.Add($row)
}

Write-Host "Done"
Write-Host -NoNewLine "Extracting from Balance Sheet................"

#Navigates to Balance Sheet

$ie.navigate($balanceurl[1])

#let page load
Start-Sleep -s 3
while($ie.busy) {Start-Sleep -s 3}
Start-Sleep -s 3

#Extracting relevant information from the Balance Sheet
# 1. Total Cash
# 2. Inventory
# 3. CAPEX
# 4. Total Debt
# 5. Total Equity

$values = 'Accounts Receivables','Accounts Payables','Inventory','Total Cash','Total Debt','Total Equity'

#Element ID of the information from the HTML code
$elements='i3','i43','i4','ttgg1','ttg5','ttg8'

for($j=0;$j -lt 6; $j++)
{
    #Create a row in the table
    $row = $financial_table.NewRow()
    $row.Millions = $values[$j];

    #Extract values for Parameter
    $str = 'data_' + $elements[$j]
    $children=$ie.document.getElementById($str).childNodes
    for($i=0;$i -lt 5;$i++)
    {
        $temp = $null;
        $temp=$children[$i].textContent
        if ($temp.startsWith('('))
        {
          $temp = $temp -replace '[()]',"";
          $temp = 0-[double]$temp;  
        } 
        if ($temp -eq '—')
        {
            $temp = 0;
        }
        $row[$year[$i]] = $temp;
    }
    #Add the row to the table
    $financial_table.Rows.Add($row)
}

Write-Host "Done"
Write-Host -NoNewLine "Extracting from Cash Flow Statement.........."

#Navigates to Cash Flow Statement

$ie.navigate($balanceurl[2])

#let page load
Start-Sleep -s 3
while($ie.busy) {Start-Sleep -s 3}
Start-Sleep -s 3

#Extracting relevant information from the Cash Flow Statement
# 1. Operating Cash Flow
# 2. CAPEX
# 3. Free Cash Flow

$values = 'Operating Cash Flow','CAPEX','Free Cash Flow'

#Element ID of the information from the HTML code
$elements='i100','i96','i97'

for($j=0;$j -lt 3; $j++)
{
    #Create a row in the table
    $row = $financial_table.NewRow()
    $row.Millions = $values[$j];

    #Extract values for Parameter
    $str = 'data_' + $elements[$j]
    $children=$ie.document.getElementById($str).childNodes
    for($i=0;$i -lt 6;$i++)
    {
        $temp = $null;
        $temp=$children[$i].textContent
        if ($temp.startsWith('('))
        {
          $temp = $temp -replace '[()]',"";
          $temp = 0-[double]$temp;  
        }
        if ($temp -eq '—')
        {
            $temp = 0;
        }
        $row[$year[$i]] = $temp;
    }
    #Add the row to the table
    $financial_table.Rows.Add($row)
}

Write-Host "Done"

$ie.quit()

#Displays data in the console for quick view
$financial_table | format-table -AutoSize

Write-Host "************************Checking with GIM Requirements (IMPORTANT!!!)************************" -ForegroundColor Red

########## Now it is time to check against GIM's requirements ##########
#1. Operating Profit Margin more than 8%                  (Score: 0.5)
#2. EBIT Growth (CAGR) more than 10% for the past 2 years (Score: 1)
#3. Net Debt to Equity less than 40%                      (Score: 1.5)
#4. Return of Equity (ROE) greater than 15%               (Score: 1.5)
#5. Quality of Earnings (QoE) greater than 80%            (Score: 1)
#6. Vampire Ratio (CAPEX/OCF) less than 50%               (Score: 1)
#7. Cash Conversion Cycle (CCC) less than 130 days        (Score: 1)
#   Total Score                                           (Score: 7.5)

$gim_table = New-Object system.Data.DataTable "$tabName";
$gim_table.Columns.Add("Criteria")| out-null;
$gim_table.Columns.Add("Value")| out-null;
$gim_table.Columns.Add("Score")| out-null;


#1. Operating Profit Margin Check
$OPM = [math]::Round(($financial_table.($year[5])[3]/$financial_table.($year[5])[0])*100,2);
if ($OPM -ge 8)
{
    $score = 0.5;
} 
else
{ 
    $score = 0;
} 

$rowcriteria = "Operating Profit Margin more than 8%";
$rowvalue = [string]($OPM) + '%';
$rowScore = [string]($score) + "/0.5";
$gim_table.Rows.Add($rowcriteria,$rowvalue,$rowscore)| out-null;
$gim_score += $score;

#2. EBIT Growth Check
$EBIT_growth = [math]::Round(([math]::sqrt($financial_table.($year[5])[3]/$financial_table.($year[3])[3])-1)*100,2);
if ([double]$financial_table.($year[5])[3] -lt 0)
#check for negative growth issues, when dividing negative with negative, it cancels out and result in false positive EBIT growth
{
    $EBIT_growth = 0-$EBIT_growth;
}
if ($EBIT_growth -ge 10) 
{
    $score = 1;
} 
else
{ 
    $score = 0;
} 

$rowcriteria = "EBIT Growth (CAGR) more than 10% for the past 2 years";
$rowvalue = [string]($EBIT_growth) + '%';
$rowScore = [string]($score) + "/1";
$gim_table.Rows.Add($rowcriteria,$rowvalue,$rowscore)| out-null;
$gim_score += $score;

#3. Net Debt to Equity Check
$gear_ratio = [math]::Round((($financial_table.($year[4])[9]-$financial_table.($year[4])[8])/$financial_table.($year[4])[10])*100,2);
if ($gear_ratio -ge 40)
{
    $score = 0;
} 
else
{ 
    $score = 1.5;
} 

$rowcriteria = "Net Debt to Equity less than 40%";
$rowvalue = [string]($gear_ratio) + '%';
$rowScore = [string]($score) + "/1.5";
$gim_table.Rows.Add($rowcriteria,$rowvalue,$rowscore)| out-null;
$gim_score += $score;

#4. Return of Equity Check
$roe = [math]::Round(($financial_table.($year[5])[3]/$financial_table.($year[4])[10])*100,2);
if ($roe -ge 15)
{
    $score = 1.5;
} 
else
{ 
    $score = 0;
} 

$rowcriteria = "Return of Equity (ROE) greater than 15%";
$rowvalue = [string]($roe) + '%';
$rowScore = [string]($score) + "/1.5";
$gim_table.Rows.Add($rowcriteria,$rowvalue,$rowscore)| out-null;
$gim_score += $score;

#5. Quality of Earnings Check
$qoe = [math]::Round(($financial_table.($year[5])[11]/$financial_table.($year[5])[3])*100,2);
if ([double]$financial_table.($year[5])[3] -lt 0)
#checks for negative operating income, which results in income loss
{
    $qoe = 0-$qoe;
}
if ($qoe -ge 80)
{
    $score = 1;
} 
else
{ 
    $score = 0;
} 

$rowcriteria = "Quality of Earnings (QoE) greater than 80%";
$rowvalue = [string]($qoe) + '%';
$rowScore = [string]($score) + "/1";
$gim_table.Rows.Add($rowcriteria,$rowvalue,$rowscore)| out-null;
$gim_score += $score;

#6. Vampire Ratio Check
$vampire_ratio = [math]::Round((0-$financial_table.($year[5])[12])/$financial_table.($year[5])[11]*100,2);
if ($vampire_ratio -ge 50 -or $vampire_ratio -lt 0)
{
    $score = 0;
} 
else
{ 
    $score = 1;
} 

$rowcriteria = "Vampire Ratio (CAPEX/OCF) less than 50%";
$rowvalue = [string]($vampire_ratio) + '%;';
$rowScore = [string]($score) + "/1";
$gim_table.Rows.Add($rowcriteria,$rowvalue,$rowscore) | out-null;
$gim_score += $score;

#7. Cash Conversion Cycle Check
$rec = [math]::Round($financial_table.($year[4])[5]/$financial_table.($year[4])[0]*365,2);
$pay = [math]::Round($financial_table.($year[4])[6]/$financial_table.($year[4])[1]*365,2);
$inv = [math]::Round($financial_table.($year[4])[7]/$financial_table.($year[4])[1]*365,2);
$ccc = $rec + $inv - $pay;
if ($ccc -lt 130)
{
    $score = 1;
} 
else
{ 
    $score = 0;
} 

$rowcriteria = "Cash Conversion Cycle (CCC) less than 130 days";
$rowvalue = [string]($ccc) + 'days';
$rowScore = [string]($score) + "/1";
$gim_table.Rows.Add($rowcriteria,$rowvalue,$rowscore) | out-null;
$gim_score += $score;

#Tabulates final score
$rowcriteria = "Total score";
$rowvalue = '-';
$rowScore = [string]($gim_score) + "/7.5";
$gim_table.Rows.Add($rowcriteria,$rowvalue,$rowscore) | out-null;

#Displays data in the console for quick view
$gim_table | format-table -AutoSize

$pwd = get-location;
$savefiledirectory = $pwd.Path + "\" + $stc;
if (!(Test-Path $savefiledirectory))
{
    mkdir $stc;
}

#Saves financial file into csv
Write-Host -NoNewLine "Saving financial file........................"
$savefilename = $stc + "\Financials_"+$stc+".csv"
$financial_table |Export-csv $savefilename
Write-Host "Done"

#Saves GIM table
Write-Host -NoNewLine "Saving GIM evaluation........................"
$savefilename = $stc + "\GIM_"+$stc+".csv"
$gim_table |Export-csv $savefilename
Write-Host "Done"
Write-Host "Your files are saved at" $savefiledirectory -ForegroundColor Green

Read-Host "Press Enter to exit"
exit
