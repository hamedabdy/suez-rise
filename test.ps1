Add-Type -AssemblyName System.Windows.Forms

$Form = New-Object system.Windows.Forms.Form
$Form.Text = "Form"
$Form.TopMost = $true
$Form.Width = 469
$Form.Height = 338

$button2 = New-Object system.windows.Forms.Button
$button2.Text = "button"
$button2.Width = 60
$button2.Height = 30
$button2.location = new-object system.drawing.point(348,249)
$button2.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($button2)

$textBox3 = New-Object system.windows.Forms.TextBox
$textBox3.Width = 159
$textBox3.Height = 20
$textBox3.location = new-object system.drawing.point(38,31)
$textBox3.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($textBox3)
Write-Host $textBox3.
$button4 = New-Object system.windows.Forms.Button
$button4.Text = "button"
$button4.Width = 60
$button4.Height = 30
$button4.location = new-object system.drawing.point(269,248)
$button4.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($button4)
Write-Host $button4

$label5 = New-Object system.windows.Forms.Label
$label5.Text = "label"
$label5.AutoSize = $true
$label5.Width = 25
$label5.Height = 10
$label5.location = new-object system.drawing.point(39,10)
$label5.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label5)

[void]$Form.ShowDialog()
$Form.Dispose()