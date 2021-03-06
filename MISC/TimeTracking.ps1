Function Set-Value {
    [cmdletbinding()]
    Param (
    $sheet,
    [string]$time, 
    [string]$day_type, 
    [string]$reason,
    [string]$place
   ) 
            $sheet.Cells.Item($j + $start_row, $time_column) = $time
            $sheet.Cells.Item($j + $start_row, $day_type_column) = $day_type
            $sheet.Cells.Item($j + $start_row, $reason_column) = $reason
			$sheet.Cells.Item($j + $start_row, $place_column) = $place
   }
   
Function AddRecipients($Mail, $EmailList, $Type)
{
    For ($i=0; $i -lt $EmailList.Length; $i++) {
        $recipientTo  = $Mail.Recipients.Add($EmailList[$i] + $postfix) 
        $recipientTo.Type = $Type; 
    }
}

Function AddContent($Mail, $Subject, $Body, $Attachment)
{
    $Mail.Subject = $Subject
    $Mail.Body =  $Body
    $rgeSource = $newSheet.range($region_copy_start, $region_copy_end).Copy()
    $Mail.Getinspector.WordEditor.Range().PasteExcelTable($true,$false,$false)
    $wdDoc = $Mail.Getinspector.WordEditor
    $wdRange = $wdDoc.Range()
    $wdRange.InsertBefore($Body)
    $Mail.Attachments.Add("$Attachment") 
}


Function OpenOutlook()
{
    #Kiểm tra Outlook đã mở chưa, nếu chưa mở thì mở cửa sổ Outlook 
    $ProcessActive = Get-Process outlook -ErrorAction SilentlyContinue
    if($ProcessActive -eq $null){
        Start Outlook -WindowStyle Maximized    
    }
    else {
        Write-host "Outlook is running"
    }

    $Outlook = New-Object -comObject  Outlook.Application 
    Return $Mail = $Outlook.CreateItem(0) 
}

Function MoveFolder($source, $destination)
{
    if((Test-Path ($destination)) -eq 0 ){
        New-Item -Path ($destination) -ItemType directory
    }
    Move-Item $source ($destination)
}

Function DeleteSheet($workBook)
{
    $i = $workBook.Worksheets.Count
    $total_sheet = $i
    While ($i -gt 1) {
        if($i -ne $total_sheet){
            $sh = $workBook.sheets.item($i)
            $sh.Delete()
        }
        $i = $i - 1

    }
}

Function SetAbsentAndLateValue
{
    [cmdletbinding()]
     Param (
        $Sheet
   ) 
    $cell_value = ""
    For ($j=0; $j -lt $total; $j++) {
        $cell_value = $Sheet.Cells.Item($j + $start_row, $email_column).Value() 
        foreach ($absent in $absent_late_list.GetEnumerator()) {
            if($cell_value  -eq $absent.Name + $postfix) {
                switch ($absent.Value) {
                    0 {   
                        Set-Value -sheet $Sheet -time $full_day -day_type $full_day_type -reason "" -place $company_place
                    }
                    1 {
                        Set-Value -sheet $Sheet -time $full_day -day_type $personal_leave -reason $absent_reason -place $company_place   
                    }
                    2 {   
                        Set-Value -sheet $Sheet -time $morning -day_type $personal_leave -reason $absent_reason -place $company_place  
                    }
                    3 {   
                        Set-Value -sheet $Sheet -time $afternoon -day_type $personal_leave -reason $absent_reason -place $company_place  
                    }
                    4 {   
                        Set-Value -sheet $Sheet -time $morning -day_type $late_type -reason $late_reason -place $company_place

                    }
    				5 {   
                        Set-Value -sheet $Sheet -time $morning -day_type $full_day_type -reason $go_out_reason -place $customer_place
                    }
    				6 { 
                        Set-Value -sheet $Sheet -time $afternoon -day_type $full_day_type -reason $go_out_reason -place $customer_place
                    }
    				7 {  
                        Set-Value -sheet $Sheet -time $full_day -day_type $full_day_type -reason $go_out_reason -place $customer_place
                    }
                }
            }
        }
    }
}

# Thư mục chứa file chấm công
$path = "E:\desktop\TaiLieuHocTap\PowerShell\BCC"
# Thư mục backup các file chấm công tháng cũ
$old_folder = "Old"

$email_list = @('anhttv14')
$cc_list = @("anhttv14")
$bcc_list = @('anhttv14')
#0 khong nghi, 1 nghi ca ngay, 2 nghi nua ngay buoi sang, 3 nghi nua ngay buoi chieu, 4 di muon, 5 khach hang buoi sang, 6 khach hang buoi chieu, 7 KH ca ngay
$absent_late_list = @{
    HungdH3 = 0
    guongvt  = 0
    Tuyetdta = 0
    anhttv14 = 0
	nganbtt = 0
    nghiepnc = 8
}

#Constant
$postfix = "@viettel.com.vn"
$region_copy_start = "A9" 
$region_copy_end = "X18"
$total = $absent_late_list.Count
$start_row = 13
$email_column = 16
$day_type_row = 9
$day_type_column = 22
$time_column = 23
$reason_column = 24
$place_column = 17
$subject = "DA EService - VIT1 gửi báo cáo quân số ngày {0}"
$body = "Hi chị,
Em gửi chị báo cáo quân số nhóm dự án Eservice ngày {0} nhé:
"
$absent_reason = "Nghỉ ốm"
$late_reason = "Muộn 2h - Hỏng xe"
$go_out_reason = "Làm việc bên khách hàng"

$full_day ="Cangay"
$morning = "Sang"
$afternoon = "Chieu"
$personal_leave = "Rv"
$late_type = "M"
$customer_place = "KH"
$company_place = "T45"
$full_day_type = "X"

#Tự động lấy ra file Excel chấm công có trong thư mục
$FilePath = Get-Item ($path + "\*") -include "*.xlsx"
$name =  [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
Write-Host $name

#Lấy ra tên file chuẩn, chỉ thay đổi tháng chấm công
$template_name = $name.Substring(0, $name.Length - 7) + "{0}.{1}"

#Lấy ra tháng và năm từ file chấm công
$month = $name.Substring($name.Length - 7, 2)
$year = $name.Substring($name.Length -4, 4)

$excel=new-object -com excel.application
$excel.Visible = $true
$workBook
$present = Get-Date

# Nếu ngày chấm công hiện tại không thuộc tháng trong file Excel
# Copy file Excel ra 1 bản mới, move file cũ vào trong thư mục OLD
# Xóa hết các sheet công để lại sheet đầu và sheet cuối
# còn nếu ngày chấm công thuộc tháng thì tiến hành copy sheet đầu tiên của file
if($month -ne $present.Month){
    $new_name = $template_name -f $present.ToString('MM.yyyy'), "xlsx"
    Copy-Item $FilePath ($path + "\" + $new_name) -Force
    
    MoveFolder -source $FilePath -destination ($path + "\" + $old_folder)
    
    $FilePath = $path + "\" + $new_name
    $workBook = $excel.workbooks.open($path + "\" + $new_name)
    $excel.DisplayAlerts = $False
    DeleteSheet -workBook $workBook
    

}
else {
    $workBook = $excel.workbooks.open($FilePath)
    $source = $workbook.Worksheets.item(1)
    $source.copy($source)
}

# Đổi tên sheet đầu tiên thành ngày chấm công hiện tại
$now= $present.ToString('dd.MM')

$newSheet = $workbook.Worksheets.Item(1)
$newSheet.Activate()
$newSheet.Name = $now

 #Chỉnh sửa giá trị ngày hiện tại
$newSheet.Cells.Item($day_type_row, $day_type_column) = "Ngày {0}" -f (Get-Date).ToString("dd/MM/yyyy") 

#Danh sách xin nghỉ, đi muộn, sang khách hàng
SetAbsentAndLateValue -Sheet $newSheet    
$workBook.Save()

Try
{
    $Mail = OpenOutlook
    AddRecipients -Mail $Mail -EmailList $email_list -Type 1
    AddRecipients -Mail $Mail -EmailList $cc_list -Type 2
    AddRecipients -Mail $Mail -EmailList $bcc_list -Type 3
    AddContent -Mail $Mail -Subject ($subject -f $now) -Body ($body -f $now) -Attachment $FilePath
    $Mail.Display()
    $Mail.Send()

    $workBook.Save()
    $excel.Workbooks.Close()
    $excel.Quit()
}
Catch
{
    Read-Host -Prompt "Complete. Press Enter to exit."
}

