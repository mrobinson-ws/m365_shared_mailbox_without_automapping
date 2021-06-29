Add-Type -AssemblyName PresentationFramework

if(-not(Get-Module ExchangeOnlineManagement -ListAvailable)){
    $null = [System.Windows.MessageBox]::Show('Please Install ExchangeOnlineManagement - view ITG for details https://worksmart.itglue.com/2503920/docs/7777752')
    Exit
}

### Start XAML and Reader to use WPF, as well as declare variables for use
[xml]$xaml = @"
<Window

  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"

  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"

  Title="J&amp;J AutoMap OUP Journals" Height="425" Width="525" ResizeMode="NoResize" SizeToContent="WidthAndHeight" Topmost="True" Background="#FFAFAFAF" Foreground="Cyan" WindowStartupLocation="CenterScreen" WindowStyle="ThreeDBorderWindow">

    <Grid Name="OKButton" Height="400" Width="519">
        <Button Name="UserButton" Content="Select User(s)" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Width="499" Height="50"/>
        <Button Name="JournalButton" Content="Select Journal(s)" HorizontalAlignment="Left" Margin="10,135,0,0" VerticalAlignment="Top" Width="499" Height="50"/>
        <Button Name="GoButton" Content="Go" HorizontalAlignment="Left" Margin="10,260,0,0" VerticalAlignment="Top" Width="499" Height="50"/>
        <RichTextBox Name="UserRichTextBox" HorizontalAlignment="Left" Height="65" Margin="10,65,0,0" VerticalAlignment="Top" Width="499" Background="#FF646464">
            <FlowDocument/>
        </RichTextBox>
        <RichTextBox Name="JournalRichTextBox" HorizontalAlignment="Left" Height="65" Margin="10,190,0,0" VerticalAlignment="Top" Width="499" Background="#FF646464">
            <FlowDocument/>
        </RichTextBox>
        <RichTextBox Name="ResultRichTextBox" HorizontalAlignment="Left" Height="65" Margin="10,315,0,0" VerticalAlignment="Top" Width="499" Background="#FF646464">
            <FlowDocument/>
        </RichTextBox>
    </Grid>

</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
Try{
    $AutoMapForm = [Windows.Markup.XamlReader]::Load($reader)
}
Catch{
    Write-Host "Unable to load Windows.Markup.XamlReader.  Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."
    Exit
}

#Create Variables For Use In Script Automatically
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {Set-Variable -Name ($_.Name) -Value $AutoMapForm.FindName($_.Name)}

# Create Functions For Color Changing Messages
Function WriteUserRichTextBox {
    Param(
        [string]$text,
        [string]$color = "Cyan"
    )

    $RichTextRange = New-Object System.Windows.Documents.TextRange( 
        $UserRichTextBox.Document.ContentEnd,$UserRichTextBox.Document.ContentEnd ) 
    $RichTextRange.Text = $text
    $RichTextRange.ApplyPropertyValue( ( [System.Windows.Documents.TextElement]::ForegroundProperty ), $color )  

}

Function WriteJournalRichTextBox {
    Param(
        [string]$text,
        [string]$color = "Cyan"
    )

    $RichTextRange = New-Object System.Windows.Documents.TextRange( 
        $JournalRichTextBox.Document.ContentEnd,$JournalRichTextBox.Document.ContentEnd ) 
    $RichTextRange.Text = $text
    $RichTextRange.ApplyPropertyValue( ( [System.Windows.Documents.TextElement]::ForegroundProperty ), $color )  

}

Function WriteResultRichTextBox {
    Param(
        [string]$text,
        [string]$color = "Cyan"
    )

    $RichTextRange = New-Object System.Windows.Documents.TextRange( 
        $ResultRichTextBox.Document.ContentEnd,$ResultRichTextBox.Document.ContentEnd ) 
    $RichTextRange.Text = $text
    $RichTextRange.ApplyPropertyValue( ( [System.Windows.Documents.TextElement]::ForegroundProperty ), $color )  

}

### End XAML, Reader, and Declarations


#Test And Connect To Microsoft Exchange Online If Needed
try {
    Get-Mailbox -ErrorAction Stop | Out-Null
}
catch {
    Connect-ExchangeOnline
}


UserButton.Add_Click({

})

JournalButton.Add_Click({

})

GoButton.Add_Click({

})

$null = $AutoMapForm.ShowDialog()