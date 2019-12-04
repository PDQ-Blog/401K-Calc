$FormatNumbers = {
    $this.Text -match '[0-9]'
    $this.Text = $this.Text -replace '[a-z]', ""
}
$AddCommas = {
    $this.Text = '{0:N0}' -f [int]$this.text
}
$CalculatePercentage = {
    if ([int]$this.Text -lt 1) {
        $this.Text = "{0:P1}" -f ([decimal]$this.Text)
    }
    elseif ([int]$this.Text -ge 1) {
        $this.Text = "{0:P1}" -f ([decimal]$this.Text / 100)
    }

}
$RemovePercentage = {
    $this.Text = $this.Text.Replace("%", "")
}

function New-RetirementGraph {
    param(
        [PSCustomObject]
        $Results
    )

    $table = @()
    $Year = (Get-Date).Year
    $RetirementNumber = $Results.RetirementExpense / $Results.WithdrawalRate
    $Age = [int]$Results.StartYear +1
    Do{
        $EmployeeContributions = [int]$Results.Salary * [decimal]$Results.EmployeeContributionPercentage
        if (($EmployeeContributions -gt 19000) -and ($Age -lt 50)) {
            $EmployeeContributions = 19000
        }
        elseif (($EmployeeContributions -gt 25000) -and ($Age -ge 50)) {
            $EmployeeContributions = 25000
        }
        if ($Results.EmployeeContributionPercentage -lt $Results.CompanyMaxContributionPercentage) {
            $CompanyContributions = $EmployeeContributions * $Results.CompanyContributionPercentage
        }
        Else {
            $CompanyContributions = [int]$Results.Salary * [decimal]$Results.CompanyMaxContributionPercentage * [Decimal]$Results.CompanyContributionPercentage
        }
        $Contributions = $EmployeeContributions + $CompanyContributions
        $zRate = [math]::pow((1 + $Results.InterestRate / $Results.NumContributionsPerYear), ($Results.NumContributionsPerYear))
        $Results.Principal = [math]::Round(([int]$Results.Principal * $zRate) + ($Contributions * ($zRate - 1) / $Results.InterestRate))
        $YearlyResults = @{
            Year                = $Year
            Age                 = $Age
            Salary              = $Results.Salary
            "401kBalance"       = $Results.Principal
            AnnualContribution  = $EmployeeContributions
            CompanyContribution = $CompanyContributions
            RetirementExpense   = "0"
        }
        $table += $YearlyResults
        $Year = $Year + 1
        $Age = [int]$Age +1
        $Results.Salary = [int]$Results.Salary * (1 + $Results.AnnualIncrease)
    }While ($Results.Principal -le $RetirementNumber)
    ###Create Draw Down Phase
    $FinalYear = $Age + 30
    $Drawdown = $Age..$FinalYear
    $xRate = [math]::pow((1 + $Results.RetirementInterestRate / 12), (12))
    Foreach($Retirement in $Drawdown){
        $Results.Principal = [math]::Round(([int]$Results.Principal * $xRate) - ($Results.RetirementExpense * ($xRate - 1) / $Results.RetirementInterestRate))
        $YearlyResults = @{
            Year                = $Year
            Age                 = $Age
            Salary              = $Results.Salary
            "401kBalance"       = $Results.Principal
            AnnualContribution  = "0"
            CompanyContribution = "0"
            RetirementExpense   = $Results.RetirementExpense
        }
        $table += $YearlyResults
        $Year = $Year + 1
        $Age = [int]$Age +1
        $Results.RetirementExpense = [int]$Results.RetirementExpense * (1 + $Results.InflationRate)    
    }



    [void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

    $chart1 = New-object System.Windows.Forms.DataVisualization.Charting.Chart
    $chart1.Width = 1800    
    $chart1.BackColor = [System.Drawing.Color]::WhiteSmoke

    # title 
    [void]$chart1.Titles.Add("Your Ending Balance is $($Results.Principal.ToString('N0'))")
    $chart1.Titles[0].Font = [System.Drawing.Font]::new("Arial", 12, [System.Drawing.FontStyle]::Bold)

    # legend 
    $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
    $legend.name = "Legend1"
    $chart1.Legends.Add($legend)

    $chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $chartarea.Name = "ChartArea1"
    $chartarea.AxisY.Title = "Balance"
    $chartarea.AxisX.Title = "Year"
    $chart1.ChartAreas.Add($chartarea)
    [void]$chart1.series.Add('balance')
    foreach ($datapoint in $table) {
        $x = $datapoint."Year"
        $y = $datapoint."401kBalance"
        [void]$chart1.Series["balance"].Points.addxy($x, $y)
    }

    # data series
    $chart1.Series["balance"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
    $chart1.Series["balance"].IsVisibleInLegend = $true
    $chart1.Series["balance"].BorderWidth = 3
    $chart1.Series["balance"].chartarea = "ChartArea1"
    $chart1.Series["balance"].Legend = "Legend1"
    $chart1.Series["balance"].color = "Blue"

    #Create Array to fill out sheet
    $FinalTable = @()
    foreach ($datapoint in $table) {
        $FinalTable += $datapoint | Select-Object @{Label = "Year"; Expression = { $_.Year } }, @{Label = "Age"; Expression = { $_.Age } }, 
        @{Label = "Your Contribution"; Expression = { $_.AnnualContribution.ToString('N0') } }, @{Label = "Employer Contribution"; Expression = { $_.CompanyContribution.ToString('N0') } },
        @{Label = "Balance"; Expression = { $_."401kBalance".ToString('N0') } }, @{Label = "Retirement Spending"; Expression = {$_.RetirementExpense.ToString('N0') } }
    }


    $FundGraph = New-Object Windows.Forms.Form
    $FundGraph.WindowState = "Maximized"
    $FundGraph.StartPosition = "Manual" 
    $FundGraph.Location = New-Object System.Drawing.Size(0, 0)
    $FundGraph.Text = "MoneyMoneyMoney....MONEY!" 
    $FundGraph.AutoSize = $true
    $FundSheet = New-Object System.Windows.Forms.DataGridView
    $FundSheet.Columns[5].AutoSizeMode("AllCells")
    $FundSheet.Height = 660
    $FundSheet.Width = 700
    $FundSheet.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $FundSheet.DefaultCellStyle.BackColor = "#f4f4f4"
    $FundSheet.BackgroundColor = $FundSheet.DefaultCellStyle.BackColor
    $FundSheet.Location = New-Object System.Drawing.Point(0, 300)
    $FundSheet.DataSource = [System.Collections.arraylist]$FinalTable
    $FundGraph.controls.AddRange(@($FundSheet, $chart1))
    $FundGraph.Refresh()
    $FundGraph.Add_Shown( { $FundGraph.Activate() }) 
    $FundGraph.ShowDialog()
}

Function New-RetirementData {
    $showGraphClicked = { 
        # Just return the object instead of doing a variable assignment and returning the variable
        $hash = [PSCustomObject]@{
            StartYear                        = $StartYear.Text
            RetirementExpense                = [int]$RetirementExpense.Text.Replace(",", "")
            Principal                        = [int]$Principal.Text.Replace(",", "")
            Salary                           = [int]$Salary.Text.Replace(",", "")
            EmployeeContributionPercentage   = [decimal]$EmployeeContributionPercentage.Text.Replace("%", "") / 100
            CompanyContributionPercentage    = [decimal]$CompanyContributionPercentage.Text.Replace("%", "") / 100
            CompanyMaxContributionPercentage = [decimal]$CompanyMaxContributionPercentage.Text.Replace("%", "") / 100
            AnnualIncrease                   = [decimal]$AnnualIncrease.Text.Replace("%", "") / 100
            InterestRate                     = [decimal]$InterestRate.Text.Replace("%", "") / 100
            NumContributionsPerYear          = $NumContributionsPerYear.Text
            InflationRate                    = [decimal]$InflationRate.Text.Replace("%", "") / 100
            WithdrawalRate                   = [decimal]$WithdrawalRate.Text.Replace("%", "") / 100
            RetirementInterestRate           = [decimal]$RetirementInterestRate.Text.Replace("%","") / 100
            FormCompleted                    = $true
        }
        New-RetirementGraph $hash
    }

    ###Create Form for input
    Add-Type -AssemblyName System.Windows.Forms

    $RetirementCalculator = New-Object system.Windows.Forms.Form
    $RetirementCalculator.ClientSize = '300,350'
    $RetirementCalculator.text = "Muh Monies"
    $RetirementCalculator.TopMost = $false

    ###Create all the boxes
    ######Create OK and Cancel Button
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(75, 315)
    $OKButton.Size = New-Object System.Drawing.Size(75, 23)
    $OKButton.Text = 'Show Graph'
    # Don't set a DialogResult to prevent the form from closing
    # $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $OKButton.Add_Click($showGraphClicked)
    $RetirementCalculator.AcceptButton = $OKButton

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(150, 315)
    $CancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $CancelButton.Text = 'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $RetirementCalculator.CancelButton = $CancelButton

    ######Create Current Age Box 
    $StartYearLabel = New-Object system.Windows.Forms.Label
    $StartYearLabel.text = "Current Age"
    $StartYearLabel.AutoSize = $true
    $StartYearLabel.width = 25
    $StartYearLabel.height = 10
    $StartYearLabel.location = New-Object System.Drawing.Point(10, 10)
    
    $StartYear = New-Object system.Windows.Forms.TextBox
    $StartYear.width = 50
    $StartYear.height = 20
    $StartYear.add_TextChanged($FormatNumbers)
    $StartYear.MaxLength = 2
    $StartYear.location = New-Object System.Drawing.Point(190, 5)

    ######Create Principal Box    
    $PrincipalLabel = New-Object system.Windows.Forms.Label
    $PrincipalLabel.text = "Current 401K Balance"
    $PrincipalLabel.AutoSize = $true
    $PrincipalLabel.width = 25
    $PrincipalLabel.height = 10
    $PrincipalLabel.location = New-Object System.Drawing.Point(10, 32)
    $Principal = New-Object system.Windows.Forms.TextBox

    $Principal.width = 50
    $Principal.height = 20
    $Principal.add_TextChanged( { $FormatNumbers })
    $Principal.add_LostFocus($AddCommas)
    $Principal.location = New-Object System.Drawing.Point(190, 27)

    ######Create Salary Box
    $SalaryLabel = New-Object System.Windows.Forms.Label
    $SalaryLabel.text = "Current Salary"
    $SalaryLabel.AutoSize = $true
    $SalaryLabel.width = 25
    $SalaryLabel.height = 10
    $SalaryLabel.location = New-Object System.Drawing.Point(10, 54)

    $Salary = New-Object System.Windows.Forms.TextBox
    $Salary.width = 50
    $Salary.height = 20
    $Salary.add_TextChanged($FormatNumbers)
    $Salary.add_LostFocus($AddCommas)
    $Salary.location = New-Object System.Drawing.Point(190, 49)

    ######Create Employee Contribution Percentage Box
    $EmployeeContributionPercentageLabel = New-Object system.Windows.Forms.Label
    $EmployeeContributionPercentageLabel.text = "401k Contribution %"
    $EmployeeContributionPercentageLabel.AutoSize = $true
    $EmployeeContributionPercentageLabel.width = 25
    $EmployeeContributionPercentageLabel.height = 10
    $EmployeeContributionPercentageLabel.location = New-Object System.Drawing.Point(10, 76)

    $EmployeeContributionPercentage = New-Object system.Windows.Forms.TextBox
    $EmployeeContributionPercentage.width = 50
    $EmployeeContributionPercentage.height = 20
    $EmployeeContributionPercentage.add_LostFocus($CalculatePercentage)
    $EmployeeContributionPercentage.add_TextChanged($FormatNumbers)
    $EmployeeContributionPercentage.add_GotFocus($RemovePercentage)
    $EmployeeContributionPercentage.location = New-Object System.Drawing.Point(190, 71)

    ######Create Employee Contribution Percentage Box
    $CompanyContributionPercentageLabel = New-Object system.Windows.Forms.Label
    $CompanyContributionPercentageLabel.text = "Employer Match%"
    $CompanyContributionPercentageLabel.AutoSize = $true
    $CompanyContributionPercentageLabel.width = 25
    $CompanyContributionPercentageLabel.height = 10
    $CompanyContributionPercentageLabel.location = New-Object System.Drawing.Point(10, 98)

    $CompanyContributionPercentage = New-Object system.Windows.Forms.TextBox
    $CompanyContributionPercentage.width = 50
    $CompanyContributionPercentage.height = 20
    $CompanyContributionPercentage.Text = "50.0%"
    $CompanyContributionPercentage.add_LostFocus($CalculatePercentage)
    $CompanyContributionPercentage.add_TextChanged($FormatNumbers)
    $CompanyContributionPercentage.add_GotFocus($RemovePercentage)
    $CompanyContributionPercentage.location = New-Object System.Drawing.Point(190, 93)

    ######Create Employee Contribution Max Percentage Box
    $CompanyMaxContributionPercentageLabel = New-Object system.Windows.Forms.Label
    $CompanyMaxContributionPercentageLabel.text = "Employer Max Match(% of Salary)"
    $CompanyMaxContributionPercentageLabel.AutoSize = $true
    $CompanyMaxContributionPercentageLabel.width = 25
    $CompanyMaxContributionPercentageLabel.height = 10
    $CompanyMaxContributionPercentageLabel.location = New-Object System.Drawing.Point(10, 120)

    $CompanyMaxContributionPercentage = New-Object system.Windows.Forms.TextBox
    $CompanyMaxContributionPercentage.width = 50
    $CompanyMaxContributionPercentage.height = 20
    $CompanyMaxContributionPercentage.add_LostFocus($CalculatePercentage)
    $CompanyMaxContributionPercentage.add_TextChanged($FormatNumbers)
    $CompanyMaxContributionPercentage.add_GotFocus($RemovePercentage)
    $CompanyMaxContributionPercentage.location = New-Object System.Drawing.Point(190, 115)

    ######Create Annual Salary Increase Box
    $AnnualIncreaseLabel = New-Object system.Windows.Forms.Label
    $AnnualIncreaseLabel.text = "Annual Salary Increase %"
    $AnnualIncreaseLabel.AutoSize = $true
    $AnnualIncreaseLabel.width = 25
    $AnnualIncreaseLabel.height = 10
    $AnnualIncreaseLabel.location = New-Object System.Drawing.Point(10, 142)

    $AnnualIncrease = New-Object system.Windows.Forms.TextBox
    $AnnualIncrease.width = 50
    $AnnualIncrease.height = 20
    $AnnualIncrease.add_LostFocus($CalculatePercentage)
    $AnnualIncrease.add_TextChanged($FormatNumbers)
    $AnnualIncrease.add_GotFocus($RemovePercentage)
    $AnnualIncrease.location = New-Object System.Drawing.Point(190, 137)

    ######Create Interest Rate Box
    $InterestRateLabel = New-Object system.Windows.Forms.Label
    $InterestRateLabel.text = "Average Rate of Return"
    $InterestRateLabel.AutoSize = $true
    $InterestRateLabel.width = 25
    $InterestRateLabel.height = 10
    $InterestRateLabel.location = New-Object System.Drawing.Point(10, 164)

    $InterestRate = New-Object system.Windows.Forms.TextBox
    $InterestRate.width = 50
    $InterestRate.height = 20
    $InterestRate.add_LostFocus($CalculatePercentage)
    $InterestRate.add_TextChanged($FormatNumbers)
    $InterestRate.add_GotFocus($RemovePercentage)
    $InterestRate.location = New-Object System.Drawing.Point(190, 159)

    ######Create Number of Contributions per year Box
    $NumContributionsPerYearLabel = New-Object system.Windows.Forms.Label
    $NumContributionsPerYearLabel.text = "Contributions Per Year"
    $NumContributionsPerYearLabel.AutoSize = $true
    $NumContributionsPerYearLabel.width = 25
    $NumContributionsPerYearLabel.height = 10
    $NumContributionsPerYearLabel.location = New-Object System.Drawing.Point(10, 186)

    $NumContributionsPerYear = New-Object system.Windows.Forms.TextBox
    $NumContributionsPerYear.multiline = $false
    $NumContributionsPerYear.text = "12"
    $NumContributionsPerYear.width = 50
    $NumContributionsPerYear.height = 20
    $NumContributionsPerYear.add_TextChanged($FormatNumbers)
    $NumContributionsPerYear.location = New-Object System.Drawing.Point(190, 181)

    ######Retirement Number
    $RetirementExpenseLabel = New-Object system.Windows.Forms.Label
    $RetirementExpenseLabel.text = "Annual Retirement Spending"
    $RetirementExpenseLabel.AutoSize = $true
    $RetirementExpenseLabel.width = 25
    $RetirementExpenseLabel.height = 10
    $RetirementExpenseLabel.location = New-Object System.Drawing.Point(10, 208)

    $RetirementExpense = New-Object system.Windows.Forms.TextBox
    $RetirementExpense.multiline = $false
    $RetirementExpense.text = ""
    $RetirementExpense.width = 50
    $RetirementExpense.height = 20
    $RetirementExpense.add_TextChanged($FormatNumbers)
    $RetirementExpense.add_LostFocus($AddCommas)
    $RetirementExpense.location = New-Object System.Drawing.Point(190, 203)

    ######Create Inflation Rate Box
    $InflationRateLabel = New-Object system.Windows.Forms.Label
    $InflationRateLabel.text = "Inflation"
    $InflationRateLabel.AutoSize = $true
    $InflationRateLabel.width = 25
    $InflationRateLabel.height = 10
    $InflationRateLabel.location = New-Object System.Drawing.Point(10, 230)

    $InflationRate = New-Object system.Windows.Forms.TextBox
    $InflationRate.width = 50
    $InflationRate.height = 20
    $InflationRate.Text = "2.0%"
    $InflationRate.add_LostFocus($CalculatePercentage)
    $InflationRate.add_TextChanged($FormatNumbers)
    $InflationRate.add_GotFocus($RemovePercentage)
    $InflationRate.location = New-Object System.Drawing.Point(190, 225)

    ######Create Withdrawal Rate Box
    $WithdrawalRateLabel = New-Object system.Windows.Forms.Label
    $WithdrawalRateLabel.text = "Planned Withdraw Rate"
    $WithdrawalRateLabel.AutoSize = $true
    $WithdrawalRateLabel.width = 25
    $WithdrawalRateLabel.height = 10
    $WithdrawalRateLabel.location = New-Object System.Drawing.Point(10, 252)

    $WithdrawalRate = New-Object system.Windows.Forms.TextBox
    $WithdrawalRate.width = 50
    $WithdrawalRate.height = 20
    $WithdrawalRate.Text = "4.0%"
    $WithdrawalRate.add_LostFocus($CalculatePercentage)
    $WithdrawalRate.add_TextChanged($FormatNumbers)
    $WithdrawalRate.add_GotFocus($RemovePercentage)
    $WithdrawalRate.location = New-Object System.Drawing.Point(190, 247)

    ######Create After Retirement Interest Rate
    $RetirementInterestRateLabel = New-Object system.Windows.Forms.Label
    $RetirementInterestRateLabel.text = "Retirement Rate of Return"
    $RetirementInterestRateLabel.AutoSize = $true
    $RetirementInterestRateLabel.width = 25
    $RetirementInterestRateLabel.height = 10
    $RetirementInterestRateLabel.location = New-Object System.Drawing.Point(10, 274)

    $RetirementInterestRate = New-Object system.Windows.Forms.TextBox
    $RetirementInterestRate.width = 50
    $RetirementInterestRate.height = 20
    $RetirementInterestRate.add_LostFocus($CalculatePercentage)
    $RetirementInterestRate.add_TextChanged($FormatNumbers)
    $RetirementInterestRate.add_GotFocus($RemovePercentage)
    $RetirementInterestRate.location = New-Object System.Drawing.Point(190, 269)

    ######Build final form
    $RetirementCalculator.controls.AddRange(@($StartYearLabel, $StartYear, $EndYearLabel, $EndYear, $PrincipalLabel, $Principal, $SalaryLabel, $Salary,
            $EmployeeContributionPercentageLabel, $EmployeeContributionPercentage, $CompanyContributionPercentageLabel, $CompanyContributionPercentage,
            $CompanyMaxContributionPercentageLabel, $CompanyMaxContributionPercentage, $AnnualIncreaseLabel, $AnnualIncrease, $InterestRateLabel, $InterestRate, $NumContributionsPerYearLabel,
            $NumContributionsPerYear, $RetirementExpenseLabel, $RetirementExpense, $InflationRateLabel, $InflationRate, $WithdrawalRateLabel, $WithdrawalRate,
            $RetirementInterestRateLabel,$RetirementInterestRate, $OKButton, $CancelButton))
    $FormResults = $RetirementCalculator.Showdialog()

    
} #End Function 


 
#Call the Function 
$Results = New-RetirementData
