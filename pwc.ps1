[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [String] $FullPath
)
# TODO: Refactor Code
class PWCExcelGenerator {
    Static [String] $MD5_Validator = "^[a-fA-F0-9]{32}$"
    Static [String] $SHA1_Validator = "^[a-fA-F0-9]{40}$"
    Static [String] $SHA256_Validator = "^[a-fA-F0-9]{64}$"
    Static [String] $DomainValidator = "^(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,}(?:\.[a-zA-Z]{2,})?$"   
    Static [String] $URLValidator = "^(https?|hxxps?|ftp):\/\/[^\s/$.?#].[^\s]*$"
    Static [String] $EmailValidator = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    Static [String] $IPV4Validator = "^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
    Static [String] $IPV6Validator = "^(([0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,7}:|([0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|([0-9a-fA-F]{1,4}:){1,5}(:[0-9a-fA-F]{1,4}){1,2}|([0-9a-fA-F]{1,4}:){1,4}(:[0-9a-fA-F]{1,4}){1,3}|([0-9a-fA-F]{1,4}:){1,3}(:[0-9a-fA-F]{1,4}){1,4}|([0-9a-fA-F]{1,4}:){1,2}(:[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:((:[0-9a-fA-F]{1,4}){1,6})|:((:[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(:[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(ffff(:0{1,4}){0,1}:){0,1}((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])|([0-9a-fA-F]{1,4}:){1,4}:((25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(25[0-5]|(2[0-4]|1{0,1}[0-9]){0,1}[0-9]))$"
    [System.Collections.Generic.List[String]] $listOfIOCs
    [System.Collections.Generic.List[String]] $MD5
    [System.Collections.Generic.List[String]] $SHA1
    [System.Collections.Generic.List[String]] $SHA256
    [System.Collections.Generic.List[String]] $Domains
    [System.Collections.Generic.List[String]] $URLS
    [System.Collections.Generic.List[String]] $IPS
    [System.Collections.Generic.List[String]] $Emails
    [System.Collections.Generic.List[String]] $OtherIOCs
    
    PWCExcelGenerator([String]$pwcFilePath) {
        $this.listOfIOCs = New-Object System.Collections.Generic.List[String]
        $this.MD5 = New-Object System.Collections.Generic.List[String]
        $this.SHA1 = New-Object System.Collections.Generic.List[String]
        $this.SHA256 = New-Object System.Collections.Generic.List[String]
        $this.Domains = New-Object System.Collections.Generic.List[String]
        $this.URLS = New-Object System.Collections.Generic.List[String]
        $this.IPS = New-Object System.Collections.Generic.List[String]
        $this.Emails = New-Object System.Collections.Generic.List[String]
        $this.OtherIOCs = New-Object System.Collections.Generic.List[String]
       
        # Extracting Values from Original PWC files
        $pwcValues = [PWCValuesExtractor]::new($pwcFilePath)
        $pwcValues.extractValues()
        $this.listOfIOCs = $pwcValues.getValues()
    }

    # Seperate/Load IOCs
    [Void] iocsExtractor() {
        forEach ($ioc in $this.listOfIOCs) {
            
            $ioc = ($ioc.ToLower()).Trim()
            
            # Removing Sanitization 
            if ($ioc.Contains("[:]")) {
                $ioc = $ioc.Replace("[:]", ":")
            }

            if ($ioc.Contains("[.]")) {
                $ioc = $ioc.Replace("[.]", ".")
            }

            #IP validation
            if ($ioc -match [PWCExcelGenerator]::IPV4Validator -or $ioc -match [PWCExcelGenerator]::IPV6Validator) {
                $this.IPS.Add($ioc)
                Continue;
            }

            #Domains validation
            if ($ioc -match [PWCExcelGenerator]::DomainValidator) {
                $this.Domains.Add($ioc)
                Continue;
            }

            #Emails validation
            if ($ioc -match [PWCExcelGenerator]::EmailValidator) {
                $this.Emails.Add($ioc)
                Continue;
            }

            #MD5 validation
            if ($ioc -match [PWCExcelGenerator]::MD5_Validator) {
                $this.MD5.Add($ioc)
                Continue;
            }

            #SHA1 validation
            if ($ioc -match [PWCExcelGenerator]::SHA1_Validator) {
                $this.SHA1.Add($ioc)
                Continue;
            }

            #SHA256 validation
            if ($ioc -match [PWCExcelGenerator]::SHA256_Validator) {
                $this.SHA256.Add($ioc)
                Continue;
            }

            #URL validation
            if ($ioc -match [PWCExcelGenerator]::URLValidator) {
                $this.URLS.Add($ioc)
                Continue;
            }
            $this.OtherIOCs.Add($ioc)
        }
        $this.displayStatus()
    }

    [Void]  displayStatus() {
        Write-Host "Total IOCs: $($this.listOfIOCs.Count)" -ForegroundColor Green
        Write-Host "MD5: $($this.MD5.Count)" -ForegroundColor Green
        Write-Host "SHA1: $($this.SHA1.Count)" -ForegroundColor Green
        Write-Host "SHA256: $($this.SHA256.Count)" -ForegroundColor Green
        Write-Host "Domains: $($this.Domains.Count)" -ForegroundColor Green
        Write-Host "URL: $($this.URLS.Count)" -ForegroundColor Green
        Write-Host "Emails: $($this.Emails.Count)" -ForegroundColor Green
        Write-Host "IPs: $($this.IPS.Count)" -ForegroundColor Green
        Write-Host "Other IOCs: $($this.OtherIOCs.Count)" -ForegroundColor Green
    }

    [Void] generateFile() {
        $excel = New-Object -ComObject Excel.Application
        $workBook = $excel.Workbooks.Add()
        $currentDateAndTime = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
        $path = (Get-Location).Path + "\PWC-$currentDateAndTime.xlsx"
        Write-Host "Creating File -- $path" -ForegroundColor Green
        $SheetCount = 2
        
        Write-Host "Adding IOCs to File -- $path. It may take some time." -ForegroundColor Yellow
        for ($i = 1; $i -le $SheetCount; $i++) {

            $workSheet = $workBook.Worksheets.Add()
            if ($i -eq 2) {
                $workSheet.Name = "PWC - IOCs"
                $Col = 0
                
                if ($this.MD5.Count -ne 0) {
                    $row = 2
                    $Col += 2
                    $cell = $workSheet.Cells[$row, $col]
                    $cell.Value = "MD5"
                    $Cell.Interior.ColorIndex = 37
                    $cell.Borders.LineStyle = 1
                    $cell.Borders.ColorIndex = 1
                    $cell.HorizontalAlignment = 3
                
                    forEach ($hash in $this.MD5) {
                        $row++
                        $cell = $workSheet.Cells[$row, $col]
                        $cell.Value = $hash
                        $cell.Borders.LineStyle = 1
                        $cell.Borders.ColorIndex = 1
                    }
                }
        
                if ($this.SHA1.Count -ne 0) {
                    $row = 2
                    $Col += 2
                    $cell = $workSheet.Cells[$row, $col]
                    $cell.Value = "SHA1"
                    $Cell.Interior.ColorIndex = 37
                    $cell.Borders.LineStyle = 1
                    $cell.Borders.ColorIndex = 1
                    $cell.HorizontalAlignment = 3
                
                    forEach ($hash in $this.SHA1) {
                        $row++
                        $cell = $workSheet.Cells[$row, $col]
                        $cell.Value = $hash
                        $cell.Borders.LineStyle = 1
                        $cell.Borders.ColorIndex = 1
                    }
                }
                
                if ($this.SHA256.Count -ne 0) {
                    $row = 2
                    $Col += 2
                    $cell = $workSheet.Cells[$row, $col]                    
                    $cell.Value = "SHA256"
                    $Cell.Interior.ColorIndex = 37
                    $cell.Borders.LineStyle = 1
                    $cell.Borders.ColorIndex = 1
                    $cell.HorizontalAlignment = 3
                
                    forEach ($hash in $this.SHA256) {
                        $row++
                        $cell = $workSheet.Cells[$row, $col]
                        $cell.Value = $hash
                        $cell.Borders.LineStyle = 1
                        $cell.Borders.ColorIndex = 1
                    }
                }
                
                if ($this.Domains.Count -ne 0) {
                    $row = 2
                    $Col += 2
                    $cell = $workSheet.Cells[$row, $col]
                    $cell.Value = "Domains"
                    $Cell.Interior.ColorIndex = 37
                    $cell.Borders.LineStyle = 1
                    $cell.Borders.ColorIndex = 1
                    $cell.HorizontalAlignment = 3
                
                    forEach ($domain in $this.Domains) {
                        $row++
                        $cell = $workSheet.Cells[$row, $col]      
                        $cell.Value = $domain
                        $cell.Borders.LineStyle = 1
                        $cell.Borders.ColorIndex = 1
                    }
                }
                
                if ($this.URLS.Count -ne 0) {
                    $row = 2
                    $Col += 2
                    $cell = $workSheet.Cells[$row, $col] 
                    $cell.Value = "URLs"
                    $Cell.Interior.ColorIndex = 37
                    $cell.Borders.LineStyle = 1
                    $cell.Borders.ColorIndex = 1
                    $cell.HorizontalAlignment = 3
                
                    forEach ($uri in $this.URLs) {
                        $row++
                        $cell = $workSheet.Cells[$row, $col] 
                        $cell.Value = $uri
                        $cell.Borders.LineStyle = 1
                        $cell.Borders.ColorIndex = 1
                    }
                }
                
                if ($this.Emails.Count -ne 0) {
                    $row = 2
                    $Col += 2
                    $cell = $workSheet.Cells[$row, $col] 
                    $cell.Value = "Emails"
                    $Cell.Interior.ColorIndex = 37
                    $cell.Borders.LineStyle = 1
                    $cell.Borders.ColorIndex = 1
                    $cell.HorizontalAlignment = 3

                    forEach ($Email in $this.Emails) {
                        $row++
                        $cell = $workSheet.Cells[$row, $col] 
                        $cell.Value = $Email
                        $cell.Borders.LineStyle = 1
                        $cell.Borders.ColorIndex = 1
                    }
                }
                
                if ($this.IPS.Count -ne 0) {
                    $row = 2
                    $Col += 2
                    $cell = $workSheet.Cells[$row, $col] 
                    $cell.Value = "IPs"
                    $Cell.Interior.ColorIndex = 37
                    $cell.Borders.LineStyle = 1
                    $cell.Borders.ColorIndex = 1
                    $cell.HorizontalAlignment = 3
                
                    forEach ($IP in $this.IPS) {
                        $row++
                        $cell = $workSheet.Cells[$row, $col] 
                        $cell.value = $IP
                        $cell.Borders.LineStyle = 1
                        $cell.Borders.ColorIndex = 1
                    }
                }
                
                #If there are Any Other values exist(s) -- For Reviewing
                if ($this.OtherIOCs.Count -ne 0) {
                    $row = 2
                    $Col += 2
                    $cell = $workSheet.Cells[$row, $col] 
                    $cell.Value = "Review -- IOC"
                    $Cell.Interior.ColorIndex = 37
                    $cell.Borders.LineStyle = 1
                    $cell.Borders.ColorIndex = 1
                    $cell.HorizontalAlignment = 3
                
                    forEach ($IOC in $this.OtherIOCs) {
                        $row++
                        $cell = $workSheet.Cells[$row, $col] 
                        $cell.Value = $IOC
                        $cell.Borders.LineStyle = 1
                        $cell.Borders.ColorIndex = 1
                    }
                }
        
            }
           
            #  Second Sheet -- Count
            if ($i -eq 1) {
                $workSheet.Name = "Count"
                $row = 2
                $col = 2
                $cell = $workSheet.Cells[$row, $col] 
                $cell.Value = "IOC Type"
                $cell.Borders.LineStyle = 1
                $cell.Borders.ColorIndex = 1
                $cell.Interior.ColorIndex = 37 
                $cell.HorizontalAlignment = 2
                
                $workSheet.Cells.Item($row, $col + 1 ) = "Count"
                $workSheet.Cells.Item($row, $col + 1).Borders.LineStyle = 1
                $workSheet.Cells.Item($row, $col + 1).Borders.ColorIndex = 1
                $workSheet.Cells.Item($row, $col + 1).Interior.ColorIndex = 37 
                $workSheet.Cells.Item($row, $col + 1).HorizontalAlignment = 3
           
                if ($this.MD5.Count -ne 0) {
                    $row += 1
                    $cell = $workSheet.Cells[$row, $col] 
                    $cell.Value = "MD5"
                    $cell.Borders.LineStyle = 1
                    $cell.Borders.ColorIndex = 1
                    $cell.HorizontalAlignment = 2
                    $workSheet.Cells.Item($row, $col + 1) = $this.MD5.Count
                    $workSheet.Cells.Item($row, $col + 1).Borders.LineStyle = 1
                    $workSheet.Cells.Item($row, $col + 1).Borders.ColorIndex = 1
                    $workSheet.Cells.Item($row, $col + 1).HorizontalAlignment = 3
                }
           
                if ($this.SHA1.Count -ne 0) {
                    $row += 1
                    $cell = $workSheet.Cells[$row, $col] 
                    $cell.Value = "SHA1"
                    $cell.Borders.LineStyle = 1
                    $cell.Borders.ColorIndex = 1
                    $cell.HorizontalAlignment = 2
                    $workSheet.Cells.Item($row, $col + 1) = $this.SHA1.Count
                    $workSheet.Cells.Item($row, $col + 1).Borders.LineStyle = 1
                    $workSheet.Cells.Item($row, $col + 1).Borders.ColorIndex = 1
                    $workSheet.Cells.Item($row, $col + 1).HorizontalAlignment = 3
                }
        
           
                if ($this.SHA256.Count -ne 0) {
                    $row += 1
                    $cell = $workSheet.Cells[$row, $col] 
                    $cell.Value = "SHA256"
                    $cell.Borders.LineStyle = 1
                    $cell.Borders.ColorIndex = 1
                    $cell.HorizontalAlignment = 2
                    $workSheet.Cells.Item($row, $col + 1) = $this.SHA256.Count
                    $workSheet.Cells.Item($row, $col + 1).Borders.LineStyle = 1
                    $workSheet.Cells.Item($row, $col + 1).Borders.ColorIndex = 1
                    $workSheet.Cells.Item($row, $col + 1).HorizontalAlignment = 3
                }
           
                if ($this.Domains.Count -ne 0) {
                    $row += 1
                    $cell = $workSheet.Cells[$row, $col] 
                    $cell.Value = "Domains"
                    $cell.Borders.LineStyle = 1
                    $cell.Borders.ColorIndex = 1
                    $cell.HorizontalAlignment = 2
                    $workSheet.Cells.Item($row, $col + 1) = $this.Domains.Count
                    $workSheet.Cells.Item($row, $col + 1).Borders.LineStyle = 1
                    $workSheet.Cells.Item($row, $col + 1).Borders.ColorIndex = 1
                    $workSheet.Cells.Item($row, $col + 1).HorizontalAlignment = 3
                }
           
                if ($this.URLS.Count -ne 0) {
                    $row += 1
                    $cell = $workSheet.Cells[$row, $col] 
                    $cell.Value = "URL"
                    $cell.Borders.LineStyle = 1
                    $cell.Borders.ColorIndex = 1
                    $cell.HorizontalAlignment = 2
                    $workSheet.Cells.Item($row, $col + 1) = $this.URLS.Count
                    $workSheet.Cells.Item($row, $col + 1).Borders.LineStyle = 1
                    $workSheet.Cells.Item($row, $col + 1).Borders.ColorIndex = 1
                    $workSheet.Cells.Item($row, $col + 1).HorizontalAlignment = 3
                }
           
                if ($this.Emails.Count -ne 0) {
                    $row += 1
                    $cell = $workSheet.Cells[$row, $col] 
                    $cell.Value = "Emails"
                    $cell.Borders.LineStyle = 1
                    $cell.Borders.ColorIndex = 1
                    $cell.HorizontalAlignment = 2
                    $workSheet.Cells.Item($row, $col + 1) = $this.Emails.Count
                    $workSheet.Cells.Item($row, $col + 1).Borders.LineStyle = 1
                    $workSheet.Cells.Item($row, $col + 1).Borders.ColorIndex = 1
                    $workSheet.Cells.Item($row, $col + 1).HorizontalAlignment = 3
                }
           
                if ($this.IPS.Count -ne 0) {
                    $row += 1
                    $cell = $workSheet.Cells[$row, $col] 
                    $cell.Value = "IPs"
                    $cell.Borders.LineStyle = 1
                    $cell.Borders.ColorIndex = 1
                    $cell.HorizontalAlignment = 2
                    $workSheet.Cells.Item($row, $col + 1) = $this.IPS.Count
                    $workSheet.Cells.Item($row, $col + 1).Borders.LineStyle = 1
                    $workSheet.Cells.Item($row, $col + 1).Borders.ColorIndex = 1
                    $workSheet.Cells.Item($row, $col + 1).HorizontalAlignment = 3
                }
           
                if ($this.OtherIOCs.Count -ne 0) {
                    $row += 1
                    $cell = $workSheet.Cells[$row, $col] 
                    $cell.Value = "Review"
                    $cell.Borders.LineStyle = 1
                    $cell.Borders.ColorIndex = 1
                    $cell.HorizontalAlignment = 2
                    $workSheet.Cells.Item($row, $col + 1) = $this.OtherIOCs.Count
                    $workSheet.Cells.Item($row, $col + 1).Borders.LineStyle = 1
                    $workSheet.Cells.Item($row, $col + 1).Borders.ColorIndex = 1
                    $workSheet.Cells.Item($row, $col + 1).HorizontalAlignment = 3
                }
           
                $row += 1
                $total = $this.MD5.Count + $this.SHA1.Count + $this.SHA256.Count + $this.Domains.Count + $this.URLS.Count + $this.Emails.Count + $this.IPS.Count + $this.OtherIOCs.Count

                $cell = $workSheet.Cells[$row, $col] 
                $cell.Value = "Total"
                $cell.Borders.LineStyle = 1
                $cell.Borders.ColorIndex = 1
                $cell.HorizontalAlignment = 2
                $cell.Interior.ColorIndex = 37 
                
                $workSheet.Cells.Item($row, $col + 1) = $total
                $workSheet.Cells.Item($row, $col + 1).Borders.LineStyle = 1
                $workSheet.Cells.Item($row, $col + 1).Borders.ColorIndex = 1
                $workSheet.Cells.Item($row, $col + 1).HorizontalAlignment = 3
            }
        }

        Write-Host "Saving & Closing the WorkBook" -ForegroundColor Green
        $workBook.SaveAs($path)
        $workbook.Close()
        $excel.Quit()
        $exitCode = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) 
        Write-Host "Closing the File: [$path] with Exit-Code: $exitCode" -ForegroundColor Yellow
    }
}
class PWCValuesExtractor {

    [String] $filePath
    [System.Collections.Generic.List[String]] $names
    [System.Collections.Generic.List[String]] $values

    PWCValuesExtractor([String] $filePath) {
        $this.filePath = $filePath
        $this.names = New-Object System.Collections.Generic.List[String]
        $this.values = New-Object System.Collections.Generic.List[String]
    }

    [Void] extractValues() {
        $excel = New-Object -ComObject Excel.Application
        $workBook = $excel.Workbooks.Open($this.filePath)

        forEach ($sheet in $workBook.Sheets) {
            $this.names.Add($sheet.Name)
        }

        forEach ($name in $this.names) {

            if ($name -eq "Legal Disclaimer" ) {
                Write-Host "Skipping sheet: [$name]" -ForegroundColor Yellow
                continue
            }

            $sheet = $workBook.Sheets.Item($name)
            Write-Host "Extracting IOCs from the Sheet: [$name]"  -ForegroundColor Green
            $rowsCount = ($sheet.UsedRange.Rows).Count
            $c = 2    

            forEach ($r in 2..$rowsCount) {
                if ($sheet.Cells.Item($r, $c).Text) {
                    $this.values.Add($sheet.Cells.Item($r, $c).Text)
                }
            }
        }

        Write-Host "Total IOCs Extracted: $($this.values.Count)" -ForegroundColor Green
        $workbook.Close()
        $excel.Quit()
        $exitCode = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) 
        Write-Host "Closing the File:[$($this.filePath)] with Exit-Code: $exitCode" -ForegroundColor Yellow
    }
    [System.Collections.Generic.List[String]] getValues() { return $this.values }
}

if ((Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Office\*\Excel")) {
    if (!(Test-Path -Path $FullPath)) {
        Write-Error "No Excel File found at the Path: [$FullPath]"
        return;
    }
    Clear-Host
    $generator = [PWCExcelGenerator]::new($FullPath)
    $generator.iocsExtractor()
    $generator.generateFile()
}
else {
    Write-Error "No Excel Module found in the System"
}