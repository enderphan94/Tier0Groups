$dateTimeFile = (get-date).ToString("yyyy MM dd")
$outFile = $($PSScriptRoot)+"\Report-$($dateTimeFile).csv"
$outFileHTML =$($PSScriptRoot)+"\Report-$($dateTimeFile).html"
$delimiter = ","

$groupsData =@(
        "Account Operators"
        "Administrators"
        "Backup Operators"
        "Domain Admins"
        "Enterprise Admins"
        "Print Operators"
        "Schema Admins" 
        "Server Operators"
    )

$members = @()

foreach($group in $groupsData){
    $member = Get-ADGroupMember -Identity $group |?{$_.objectClass -eq "user"} | select -ExpandProperty sAMAccountName
    $members += $member
}

$userCount= $members.Count

$NeverExpires = 9223372036854775807

$global:ea = 0
$global:last2015 = 0
$global:last2016 = 0
$global:last2017 = 0
$global:otherLast = 0
$global:NeverLogon = 0
$global:noLastSet = 0
$global:passSet2015 = 0
$global:passSet2016 = 0
$global:passSet2017 = 0
$global:otherPassSet = 0
$global:accNotEx = 0
$global:accDisStatus=0
$global:smartRe =0
$global:passNExpSet = 0

foreach($member in $members){
    
    try{
        $memberInfor = Get-ADUser -Identity $member -Properties *|select distinguishedName,
                                                                    sAMAccountName,
                                                                    mail,
                                                                    lastLogonTimeStamp,
                                                                    pwdLastSet,
                                                                    badpwdcount,
                                                                    accountExpires,
                                                                    userAccountControl,
                                                                    modifyTimeStamp,
                                                                    lockoutTime,
                                                                    badPasswordTime,
                                                                    maxPwdAge,
                                                                    Description

       
        #Last Logon
        $lastLogon = [datetime]::fromfiletime($memberInfor.lastLogonTimeStam)        
        $lastLogon= $lastLogon.ToString("yyyy/MM/dd")     
        if($lastLogon.split("/")[0] -eq 2015){
            $global:last2015++
        }     
        elseif ($lastLogon.split("/")[0] -eq 2016){
            $global:last2016++
        }
        elseif ($lastLogon.split("/")[0] -eq 2017){
            $global:last2017++
        }elseif ($lastLogon.split("/")[0] -eq 1601){
            $lastLogon = "Never"
            $global:NeverLogon++
        }else{
            $global:otherLast++
        }
               
        #password last set
        if($memberInfor.pwdLastSet -eq 0)
        {         
             $pwdLastSet = "Never"
             $global:noLastSet++
        }
        else
        {         
             $pwdLastSet = [datetime]::fromfiletime($memberInfor.pwdLastSet)                   
             $pwdLastSet = $pwdLastSet.ToString("yyyy/MM/dd")
             if($pwdLastSet.split("/")[0] -eq 2015){ 
                 $global:passSet2015++
             }     
             elseif ($pwdLastSet.split("/")[0] -eq 2016){
                 $global:passSet2016++
             }
             elseif ($pwdLastSet.split("/")[0] -eq 2017){
                 $global:passSet2017++
             }
             elseif ($pwdLastSet.split("/")[0] -eq 1601){
                 $pwdLastSet = "Never"   
                 $global:noLastSet++ 
             }
             else{
                 $global:otherPassSet++
             }
         
        }     
        #Account expires   
        if(($memberInfor.accountExpires -eq $NeverExpires) -or ($memberInfor.accountExpire -gt [Datetime]::MaxValue.Ticks))
        {
            $convertAccountEx = "Not Expired"
        
        }
        else
        {
            $convertAccountEx = "Expired"
            $global:accEx++
        }
        #Email
        if([String]::IsNullOrEmpty($mail)){        
            $email = "N/A"        
        }
        else{
            $email =$mail
            $global:ea++
        }
  
        #UserInfor
        if($memberInfor.userAccountControl -band 0x0002)
        {
            $accountDisStatus = "disabled"
            $global:accDisStatus++
        }
        else
        {
            $accountDisStatus = "enabled"
        }  
        #If Smartcard Required
        if( $memberInfor.userAccountControl -band 262144)
        {
            $smartCDStatus = "Required"
            $global:smartRe++
        }
        else
        {
            $smartCDStatus = "Not Required"
        }  

        #If No password is required
        if( $memberInfor.userAccountControl -band 32){
            $passwordEnforced ="Not Required"
            $global:passNotRe++
        }
        else
        {
            $passwordEnforced = "Required"
        }  

        #Password never expired
        if( $memberInfor.userAccountControl -band 0x10000){
            $passNExp ="Never Expires is set"
            $global:passNExpSet++
        
        }
        else
        {
            $passNExp = "None Set"
            $passTrue = $true
        }  
       	  
        
        $obj = New-object -TypeName psobject
        $obj | Add-Member -MemberType NoteProperty -Name "Distinguished Name" -Value $memberInfor.distinguishedName
        $obj | Add-Member -MemberType NoteProperty -Name "Sam account" -Value $memberInfor.sAMAccountName
        $obj | Add-Member -MemberType NoteProperty -Name "Email" -Value $memberInfor.mail
        $obj | Add-Member -MemberType NoteProperty -Name "Password last changed" -Value $pwdLastSet 
        $obj | Add-Member -MemberType NoteProperty -Name "Last Logon " -Value $lastLogon
        $obj | Add-Member -MemberType NoteProperty -Name "Account Expires" -Value $convertAccountEx
        $obj | Add-Member -MemberType NoteProperty -Name "Account Status" -Value $accountDisStatus  
        $obj | Add-Member -MemberType NoteProperty -Name "Smartcard Required" -Value $smartCDStatus 
        $obj | Add-Member -MemberType NoteProperty -Name "Password Required" -Value $passwordEnforced  
        $obj | Add-Member -MemberType NoteProperty -Name "Never Expired Password Set" -Value $passNExp   
        $obj | Add-Member -MemberType NoteProperty -Name "Description" -Value $memberInfor.Description 
        #$obj
        $obj | Export-Csv -Path "$outFile" -NoTypeInformation -append -Delimiter $delimiter 
    }
    catch{
        Write-Error "Can't get this user $member"
    }

   
                                                                    
}
# Saving images to import later to the HTML reports
$global:IncludeImages = New-Object System.Collections.ArrayList 
$global:check= 0
$global:outFilePicPie = $($PSScriptRoot)+"\Pie-$($dateTimeFile)-$($global:check).jpeg"
#PIE
#Email
$emailPer = $global:ea 
$noEmailPer=  $userCount - $emailPer
$mailHash = @{"Available"=$emailPer;"Unavailable"=$noEmailPer}

#Account expired
$accExPer = $global:accEx
$accNotExPer = $userCount - $accExPer
$accExHash = @{"Expired"="$accExPer";"Unexpired"="$accNotExPer"}

#Account Status
$accDisPer = $global:accDisStatus
$accNoDisPer = $userCount - $accDisPer
$accStatusHash = @{"Disabled"="$accDisPer";"Enabled"="$accNoDisPer"}

#Smart Card required
$smartRePer = $global:smartRe
$smartNotRePer = $userCount - $smartRePer
$smartReHash = @{"Required"="$smartRePer";"Not Required"="$smartNotRePer"}

#Password Required
$passReNotPer = $global:passNotRe 
$passRePer =  $userCount - $passReNotPer
$passReHash = @{"Not Required"="$passReNotPer";"Required"="$passRePer"}

#Password Never Expired Set
$passExpSetPer =$global:passNExpSet
$passExpNoSetPer= $userCount - $passExpSetPer
$passExpHash = @{"Set"="$passExpSetPer";"None-set"="$passExpNoSetPer"}
Function drawPie {
    param($hash,
    [string]$title
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Windows.Forms.DataVisualization
    $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
    $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $Series = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.Series
    $ChartTypes = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]
    $Series.ChartType = $ChartTypes::Pie
    $Chart.Series.Add($Series)
    $Chart.ChartAreas.Add($ChartArea)
    $Chart.Series['Series1'].Points.DataBindXY($hash.keys, $hash.values)
    $Chart.Series[‘Series1’][‘PieLabelStyle’] = ‘Disabled’
    $Legend = New-Object System.Windows.Forms.DataVisualization.Charting.Legend
    $Legend.IsEquallySpacedItems = $True
    $Legend.BorderColor = 'Black'
    $Chart.Legends.Add($Legend)
    $chart.Series["Series1"].LegendText = "#VALX (#VALY)"
    $Chart.Width = 700
    $Chart.Height = 400
    $Chart.Left = 10
    $Chart.Top = 10
    $Chart.BackColor = [System.Drawing.Color]::White
    $Chart.BorderColor = 'Black'
    $Chart.BorderDashStyle = 'Solid'
    $ChartTitle = New-Object System.Windows.Forms.DataVisualization.Charting.Title
    $ChartTitle.Text = $title
    $Font = New-Object System.Drawing.Font @('Microsoft Sans Serif','12', [System.Drawing.FontStyle]::Bold)
    $ChartTitle.Font =$Font
    $Chart.Titles.Add($ChartTitle)
    $testPath = Test-Path $global:outFilePicPie
    if($testPath -eq $True){
        $global:check += 1      
        $global:outFilePicPie = $($PSScriptRoot)+"\Pie-$($dateTimeFile)-$($global:check).jpeg"                 
    }
    $global:IncludeImages.Add($global:outFilePicPie)
    $Chart.SaveImage($outFilePicPie, 'jpeg')  
}
#BAR
#lastLogon
$lastLogonHash = [ordered]@{"Never"="$global:NeverLogon";"<2015"="$global:otherLast";"2015"="$global:last2015";"2016"="$global:last2016";"2017"="$global:last2017"}
$global:check1= 0
$global:outFilePicBar = $($PSScriptRoot)+"\Bar-$($dateTimeFile)-$($global:check).jpeg"

#PassLastSet
$passSetHash = [ordered]@{"Never"="$global:noLastSet";"<2015"="$global:otherPassSet";"2015"="$global:passSet2015";
                        "2016"="$global:passSet2016";"2017"="$global:passSet2017";}

function drawBar{
    param(
    $hash,[string]$title,[string]$axisX
    ) 
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Windows.Forms.DataVisualization
    $Chart1 = New-object System.Windows.Forms.DataVisualization.Charting.Chart
    $ChartArea1 = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $Series1 = New-Object -TypeName System.Windows.Forms.DataVisualization.Charting.Series
    $ChartTypes1 = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]
    $Chart1.Series.Add($Series1)
    $Chart1.ChartAreas.Add($ChartArea1)
    $Chart1.Series[‘Series1’].Points.DataBindXY($hash.keys, $hash.values)
    $chart1.Series[0].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Column
    $ChartArea1.AxisX.Title = $axisX
    $ChartArea1.AxisY.Title = "Figures"
    $Chart1.Series[‘Series1’].IsValueShownAsLabel = $True
    $Chart1.Series[‘Series1’].SmartLabelStyle.Enabled = $True
    $chart1.Series[‘Series1’]["LabelStyle"] = "TopLeft"

    $ChartArea1.AxisY.Maximum = $userCount
    
    if($userCount -ge 1000){
        $ChartArea1.AxisY.Interval = $inter - ($inter %100)
        $inter = [math]::Round($userCount/10,0)
    }elseif($userCount -ge 100){
        $ChartArea1.AxisY.Interval = $inter - ($inter %10)
        $inter = [math]::Round($userCount/20,0)
    }else{
        $ChartArea1.AxisY.Interval = $inter - ($inter %10)
        $inter = [math]::Round($userCount/10,0)
    }    
    
    $Chart1.Width = 1000
    $Chart1.Height = 700
    $Chart1.Left = 10
    $Chart1.Top = 10
    $Chart1.BackColor = [System.Drawing.Color]::White
    $Chart1.BorderColor = 'Black'
    $Chart1.BorderDashStyle = 'Solid'      
    $ChartTitle1 = New-Object System.Windows.Forms.DataVisualization.Charting.Title
    $ChartTitle1.Text = $title
    $Font1 = New-Object System.Drawing.Font @('Microsoft Sans Serif','12', [System.Drawing.FontStyle]::Bold)
    $ChartTitle1.Font =$Font1
    $Chart1.Titles.Add($ChartTitle1)

    $testPath = Test-Path $global:outFilePicBar
    if($testPath -eq $True){
        $global:check1 += 1      
        $global:outFilePicBar = $($PSScriptRoot)+"\Bar-$($dateTimeFile)-$($global:check1).jpeg"         
    }
    $global:IncludeImages.Add($global:outFilePicBar)
    $Chart1.SaveImage("$outFilePicBar", 'jpeg')
}
# Draw Pie
drawPie -hash $mailHash -title "Emails" |Out-Null
drawPie -hash $accExHash -title "Expired Accounts"|Out-Null
drawPie -hash $accStatusHash -title "Account Status"|Out-Null
drawPie -hash $smartReHash -title "Smart Cards Required"|Out-Null
drawPie -hash $passReHash -title "Password Required"|Out-Null
drawPie -hash $passExpHash -title "Password Never Expired Settings"|Out-Null

# Draw bar
drawBar -hash $lastLogonHash -title  "Last Logon Date" -axisX "years"|Out-Null
drawBar -hash $passSetHash -title "Password Last Changed" -axisX "years"|Out-Null


# Generating HTML Reports
$userName = Read-Host "Reporter"
$groupsData = $groupsData|Out-String
$body =@'
<h1> Forest Report </h1>
<p><ins><b>I.<b> Information<ins></p>
<div class="tablehere">
    <table class="tabinfo" > 
          <tr>
            <td>Reported by:</td>
            <td>{0}</td> 
          </tr>
          <tr>
            <td>Datetime:</td>
            <td>{1}</td> 
          </tr>           
    </table>
</div>
<div class="tabofexecu">
    <table class="tabexecu" > 
          <tr>
            <td>List of tier0 groups:</td>
            <td>{2}</td> 
          </tr>
          <tr>
            <td>Number of accounts: </td>
            <td>{3}</td> 
          </tr>  
                
    </table>
<div>

<p><ins><b>III.<b> Data Illustration<ins></p>
'@ -f  $userName,$(get-date),$groupsData,$members.Count

function Generate-Html {
    Param(
        [Parameter()]
        [string[]]$IncludeImages
    )

    if ($IncludeImages){
        $ImageHTML = $IncludeImages | % {
        $ImageBits = [Convert]::ToBase64String((Get-Content $_ -Encoding Byte))
        "<center><img src=data:image/jpeg;base64,$($ImageBits) alt='My Image'/><center>"
    }
        ConvertTo-Html -Body $body -PreContent $imageHTML -Title "Report on $Domain" -CssUri "style.css" |
        Out-File $outFileHTML
    }
}

Generate-Html -IncludeImages $global:IncludeImages


foreach($image in $IncludeImages){

    rm $image 
}
#Finish
Write-Host
Write-Verbose -Message  "Script Finished!!" -Verbose
