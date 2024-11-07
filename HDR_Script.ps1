Function openXMLFile ($xmlFile) {
    try {
        #[xml]$xml = Get-Content "$xmlFile" -ErrorAction Stop
        [xml]$xml = @"
        <config>
            <datSeparators>
                <input>
                    <!-- DAT inputs could have variable field and text separators -->
                    <field>
                        <obj charName="Device Control Four" method="unicode" chr="U+0014" powershell="0x14" />
                    </field>
                    <text>
                        <obj charName="Latin Small Letter Thron (UTF8)" method="unicode" chr="U+00FE"
                            powershell="254" />
                        <obj charName="Latin Small Letter Thron (ASCII)" method="unicode" chr="U+00FE"
                            powershell="65533" />
                    </text>
                </input>

                <!-- Output how every DAT file will be outputted -->
                <output>
                    <field>
                        <obj charName="Device Control Four" method="unicode" chr="U+0014" powershell="0x14" />
                    </field>
                    <text>
                        <obj charName="Latin Small Letter Thron" method="unicode" chr="U+00FE"
                            powershell="254" />
                    </text>
                </output>

                <!-- each eDiscovery software will have a different way multivalue separator -->
                <multiValueSeparator>
                    <obj software="casepoint" charName="Right-Pointing Double Angle Quotation Mark"
                        method="unicode" chr="U+00BB" powershell="187" />
                </multiValueSeparator>
            </datSeparators>

            <!-- workflow-->
            <casepoint>
                <step id="00" name="Setup File Delimiters and Review" module="fileReviewAndSetup" />
                <step id="01" name="Setup Field Remapping" module="fieldHeaderRemappingInitialize" />
                <step id="02" name="File Remapping" module="fileRemapping" />
                <!--<
                step id="01" name="Setup File Delimiters" module="fileDelimiterSetup" /> -->
            </casepoint>
        </config>
"@
    }
    catch {
        $except = $_.Exception.Message
    }

    if ($except) {
        write-host ""
        write-host "!ERROR!" -foreground red
        write-host "Config not found, please attempt to re-run or reach out to sysAdmins"
        write-host ""
        Read-Host  "Press any key to exit"
        exit
    }
    return $xml
}

Function getDBObjects {
    $dbConn = New-Object System.Data.Odbc.OdbcConnection
    $dbConnectionString = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + $hSetup.dbPath + ";Mode=Read"
    $dbConn.ConnectionString = $dbConnectionString
    $dbConn.Open()

    Function executeCMD ($sqlQuery) {
        $dbCMD = $dbConn.CreateCommand()
        $dbCMD.CommandText = $sqlQuery

        $rdr = $dbCMD.ExecuteReader()
        $tbl = New-Object Data.DataTable
        $tbl.Load($rdr)
        $rdr.Close()
        return $tbl
    }

    # get default mapping
    $sqlObjects = @()
    $sqlObjects += "mapDefault"
    $sqlObjects += "mapOverride"
    $sqlObjects += "mapRegex"
    $sqlObjects += "softwareFieldInfo"
    $sqlObjects += "appendExtraFields"

    # setup the SQL queries.
    $sqlQuerys = @()
    foreach ($sqlObject in $sqlObjects) {
        <# $sqlObject is $sqlObjects item #>
        $sqlQuerys += "SELECT * FROM " + $hSetup.ediscoverySoftware + '_' + $sqlObject
    }

    # execute sqlQuerys
    $tempCounter = 0
    foreach ($sqlObject in $sqlObjects) {
        $returned = executeCMD $sqlQuerys[$tempCounter]
        $tempCounter++

        # convert some tables to dictionaries for faster lookup.
        if ($sqlObject -match "mapDefault|softwareFieldInfo") {
            if ($sqlObject -match "mapDefault") {
                $key = "datField"
            }
            if ($sqlObject -match "softwareFieldInfo") {
                $key = "field"
            }

            $hTemp = [ordered]@{}
            $returned | ForEach-Object {
                $hTemp.Add($_."$key", $_)
            }
            $hSetup.(("dbObject" + '_' + $sqlObject)) = $hTemp
        }
        else {
            $hSetup.(("dbObject" + '_' + $sqlObject)) = $returned
        }
    }
}

# Get total lines in file
Function totalLineNubersInCSV ($file) {
    [Linq.Enumerable]::Count([System.IO.File]::ReadAllLines($file))
}

# slightly alters and ccleans field name
Function compactHeaderFieldToRemoveExtraItem ($field) {
    # uppercase
    # trim 
    # remove space and _
    $field.toUpper().Trim() -replace " ", "" -replace "_"
}

Function openExcelObject {
    # Launch Excel com
    try {
        $excelObj = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
    }
    catch {
        $exception = $_.Exception
    }
    if ($exception) {
        Write-Host "!!ERROR!!" -Foreground Red
        Write-Host "EXCEL Application Does Not Exist"
        Write-Host "Please run the script on a machine that has EXCEL Application installed on it."
        Write-Host ""
        Read-Host  "Press Any Key To Exit."
        exit
    }

    $excelObj.Visible = $false
    $excelObj.DisplayAlerts = $false

}

Function closeExcelObject {
    # Save and Close workbook
    $workBook.Saveas(($hSetup.fileOutpath + '\' + $hSetup."mappingInfoXLSXNameOutName"))
    $excelObj.quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelObj) | Out-Null
}

# this will out field examples from the dat file.
Function outputFieldExamplesToExcel($hFieldExamples, $numOfExamples, $hitsFound, $worksheetName, $quitOrContinue) {
    # pad out empty entries if no values were found
    $hFieldExamplesFinal = [ordered]@{}
    $hFieldExamples.getEnumerator() | ForEach-Object {
        $key = $_.key
        $value = $_.value
        $numOfExamplesFound = ($value | Measure-Object).Count
        $numOfRemainExamples = $numOfExamples - $numOfExamplesFound
        if ($numOfRemainExamples -ne 0) {
            if ($numOfRemainExamples -eq $numOfExamples) {
                $value += "FIELD IS EMPTY"
            }
            else {
                foreach ($numOfExample in (1..$numOfRemainExamples)) {
                    <# $numOfExample is (1..$numOfRemainExamples) item #>
                    $value += "NO ADDITIONAL EXAMPLES"
                }
            }
        }
        $hFieldExamplesFinal.($key) = $value
    }
    # Update Excel with Header Samples
    # Access the correct worksheet
    $workSheet = $workBook.sheets.item($worksheetName)
    $rowCounter = 2
    $hFieldExamplesFinal.GetEnumerator() | ForEach-Object {
        $key = $_.Key

        $fileHeaderUnAlteredHeaders = $hSetup."fileHeader".($key) -replace "-HEADER-NORM-DUPE-FIELD\d+$"
        $colCounters = 1..3

        foreach ($el in $_.Value) {
            <# $el is t_.Value item #>
            foreach ($colCounter in $colCounters) {
                <# $colCounter is tcol$colCounters item #>
                if ($colCounter -eq 1) {
                    $rowValue = $fileHeaderUnAlteredHeaders
                }
                if ($colCounter -eq 2) {
                    $rowValue = $key
                }
                if ($colCounter -eq 2) {
                    $rowValue = $el
                }
                $workSheet.cells.item($rowCounter, $colCounter) = $rowValue
            }   
            $rowCounter++
        }
    }

    if ($quitOrContinue -match "quit") {
        # FUNCTION Excel Object
        closeExcelObject

        write-host "DAT HEADER MAPPING ISSUE" -foreground red
        write-host ""
        write-host "Source DAT:"
        write-host $hSetup.("filePath")
        write-host ""
        write-host "The DAT fields below are not configured for mapping"
        write-host ""
        write-host "Review the EXCEL document for mapping and example data."
        write-host ""
        write-host "EXCEL Location: " -nonewline; write-host ('\' + $hSetup.("fileOutFolderName") + '\' + $hSetup.("mappingInfoXLSXName")) -foreground green
        write-host "Worksheet: Mapped Header"
        write-host "Worksheet: Unmapped Header"
        write-host ""
        write-host "Please reach out to " -nonewline; write-host "USER SUPPORT" -nonewline -foreground cyan; write-host " for help mapping these new fields."
        write-host ""
        $results = $hitsFound | ForEach-Object {
            [PSCustomObject]@{
                'DAT FIELD'    = $_.name
                'MAP TO FIELD' = $_.value
            }
        }
        $results | Out-String | Format-Table -AutoSize

        if (!$hSetup.("noPrompts")) {
            Read-Host "Press any key to exit"
        }
        exit
    }
}

# Output file information to the excel 
Function outputLoadReadyFileInfo($worksheetName) {
    $fileOutputTotalNumOfLines = totalLineNubersInCSV($hSetup.fileOutpath + '\' + $hSetup.fileName)
    
    # Access the correct worksheet 
    $workSheet = $workBook.sheets.item($worksheetName)

    # Populate all the fields (this is not automated)

    # SOURCE 
    $workSheet.cells.item(2, 2) = $hSetup."folderPath"
    $workSheet.cells.item(3, 2) = $hSetup."fileName"
    $workSheet.cells.item(4, 2) = $hSetup."fileTotalLines"
    $workSheet.cells.item(5, 2) = $hSetup."filePath"

    # TARGET
    $workSheet.cells.item(7, 2) = $hSetup.fileOutpath
    $workSheet.cells.item(8, 2) = $hSetup.fileName
    $workSheet.cells.item(9, 2) = $fileOutputTotalNumOfLines
    $workSheet.cells.item(10, 2) = ($hSetup.fileOutpath + '\' + $hSetup.fileName)

    $lineColor = "green"
    $hSetup."fileAndTargetLineCountsMatch" = "true"
    if ($hSetup."fileTotalLines" -ne $fileOutputTotalNumOfLines) {
        $lineColor = "red"
        $hSetup."fileAndTargetLineCountsMatch" = "false"
    }

    write-host ""
    write-host "SOURCE LR TOTAL LINES: " -foreground $lineColor -nonewline; write-host $hSetup."fileTotalLines"
    write-host "TARGET LR TOTAL LINES: " -foreground $lineColor -nonewline; write-host $fileOutputTotalNumOfLines
    write-host ""

    if ($lineColor -eq "red") {
        write-host "Troubleshooting:"
        write-host "   1. Manually check the line counts for the SOURCE/TARGET."
        write-host "   2. If line counts are incorrect, attempt to reprocess."
        write-host "   3. If line counts are still incorrect, escalate to IRT."
        write-host ""
    }

}

# FUNCTION - this open the INPUT file to determine what field delimiters are being used.
Function fileReviewAndSetup {
    Function checkForSeparator($firstLine, $sepItem) {
        $delimiter = ""
        foreach ($el in $sepItem) {
            [char][int]$delim = $el.powershell
            #$firstline
            if ($firstLine -match $delim) {
                $delimiter = $delim
                break
            }
        }
        # if delimiter couldn't be identified, there is no point to move on
        if (!$delimiter) {
            # FUNCTION Excel Object
            closeExcelObject

            Write-Host "-------------------------" -ForegroundColor Red
            Write-Host "Input Delimiters could not be identified, please review."
            Write-Host "`nScript exiting, please fix module`n"
            Read-Host  "Press any key to exit"
            Exit
            $aGlobalHeaderIssues += $hSetup."filePath"
            break
        }

        return $delimiter

    }

    try {
        # Get total lines in file
        $hSetup."fileTotalLines" = totalLineNubersInCSV $hSetup."filePath"

        $bom = New-Object -TypeName System.Byte[](4)
        $file = New-Object System.IO.FileStream($hSetup."filePath", 'Open', 'Read')
        $null = $file.Read($bom, 0, 4)
        $file.Close()
        $file.Dispose()
        $hSetup.("fileEncoding") = "UTF8"
        if ($bom[0] -eq 0xfe -and $bom[1] -eq 0x00 -and $bom[2] -eq 0x46 -and $bom[3] -eq 0x00) {
            $hSetup."fileEncoding" = "UNICODE"
        }
        $firstLine = Get-Content $hSetup."filePath" -First 1 -encoding $hSetup."fileEncoding" -ErrorAction Stop
    }
    catch {
        <#Do this if a terminating exception happens#>
        $except = $_.Exception
        if ($except) {
            # FUNCTION Excel Object
            closeExcelObject

            Write-Host "UNABLE TO GET FIRST ROW OF DAT FILE"
            Write-Host "-------------------------" -ForegroundColor Red
            Write-Host "Input Delimiters could not be identified, please review."
            Write-Host "`nScript exiting, please fix module`n"
            Read-Host  "Press any key to exit"
            break
            exit

        }
    }

    # INPUT
    # now we retrived the all input DAT separators
    $xmlSelected = $hSetup."scriptConfigXML".SelectNodes("//datseparators/input")
    $aTemp = @()
    # Check FIELD separator break once found.
    $aTemp += checkForSeparator $firstLine $xmlSelected.field.obj
    # check TEXT separator break once found.
    $aTemp += checkForSeparator $firstLine $xmlSelected.text.obj

    # add input file field and text separator to main dict.
    $hSetup."fileDelimiterInput" = $aTemp

    # OUTPUT
    # add input file field and text separator to main dict.
    $aTemp = @()
    $xmlSelected = $hSetup."scriptConfigXML".SelectNodes("//datseparators/output")
    # FIELD separator.
    $aTemp += [char][int]$xmlSelected.field.obj.powershell
    # TEXT separator.
    $aTemp += [char][int]$xmlSelected.text.obj.powershell

    # add output file field and text separator to main dict.
    $hSetup."fileDelimiterOutput" = $aTemp

    # setup parse out first line along with apply field overrides.
    $fieldOverrides = @{}
    if ($hSetup.fieldOverride) {
        $hSetup.fieldOverride -split "," | ForEach-Object {
            $split = $_ -split "="
            $fieldOverrides.($split[0].Trim()) = $split[1].Trim()
        }
    }

    $hFileHeader = [ordered]@{}
    $headerNULLValues = 0
    $firstLine -split $hSetup."fileDelimiterInput"[0] -replace $hSetup."fileDelimiterInput"[1] | ForEach-Object {
        # check if header value is empty.
        if (!$_) {
            $headerNULLValues++
        }
        $headerCol = $_
        $compactHeaderCol = compactHeaderFieldToRemoveExtraItem $headerCol

        # when duplicate header values are found, we add a number to the end so it allows the tool to continue.
        $headerFound = $hFileHeader.($compactHeaderCol)
        if ($headerFound) {
            $headerFoundAddNum = ($headerFound -split "\|" | Measure-Object).Count + 1
            $hFileHeader[($compactHeaderCol + "-HEADER-NORM-DUPE-FIELD" + $headerFoundAddNum)] = $headerCol
        }
        else {
            $hFileHeader[$compactHeaderCol] = $headerCol
        }
    }
    if ($headerNULLValues -ne 0) {
        # FUNCTION Excel Object
        closeExcelObject

        Write-Host "!!ERROR: HEADER CONTAINS EMPTY VALUES"
        Write-Host "-------------------------" -ForegroundColor Red
        Write-Host "$name`n"
        Write-Host "ERROR: Header Contains" $headerNULLValues "Empty Fields" -ForegroundColor Red
        Write-Host "`nScript exiting, please fix header`n"
        Read-Host  "Press any key to exit"
        break
        exit
    }
    
    # perform a duplicate load ready header check. Tool will exit if found.

    # add header to setup hash
    $hSetup."fileHeader" = $hFileHeader

    # Convert header to include text delimiters
    $aHeader = @()
    $hSetup.fileHeader.Keys | ForEach-Object { $aHeader += ($hSetup.fileDelimiterInput[1] + $_ + $hSetup.fileDelimiterInput[1]) }
    $hSetup["fileHeaderForImporting"] = $aHeader

    # Are additional fields going to be append ? if yes, we check and add that to hSetup here.
    # CASEPOINT SPECIFIC
    if ($hSetup.dbObject_appendExtraFields) {
        # for this custom piece, we need to extract specific client info.
        $folderPath = $hSetup."folderPath" -replace "(^.+\\\d{4}-\d{2}-\d{2}\\\d{6})\\.+$", "`$1"
        $folderPathSplit = $folderPath -split "\\"
        $folderPathSplitSelection = $folderPathSplit | Select-Object -Last 3
        $hTemp = [ordered]@{}
        $hSetup.dbObject_appendExtraFields | ForEach-Object {
            $el = $_."mapField"
            switch ($el) {
                "Producing Party_CF" { $hTemp.($el) = $folderPathSplitSelection[0] }
                "Production Date" { $hTemp.($el) = ([datetime]$folderPathSplitSelection[1]).ToString("MM/dd/yyyy") }
                "Request ID_CF" { $hTemp.($el) = $folderPathSplitSelection[2] }
                "Dateloaded_CF" { $hTemp.($el) = $hSetup."todayDate" }
                Default {}
            }
        }
        $hSetup.("appendExtraFieldDATA") = $hTemp
    }
}

# FUNCTION - this takes the header row and compares it against the normalized header to generate the cross refrence,
Function fieldHeaderRemappingInitialize {
    # Setup parse out first line along with apply field overrides.
    $fieldOverrides = @{}
    if ($hSetup.fieldOverride) {
        $hSetup.fieldOverride -split "," | ForEach-Object {
            $split = $_ -split "="
            $fieldOverrides.($split[0].Trim()) = $split[1].Trim()
        }
    }

    # input file row counter 
    $rowCounter = 0

    [string]$ifNoFieldIsFound = "_NO_FIELD_FOUND"
    $hNewFieldMapping = [ordered]@{}
    $tempCounter = 1
    $hSetup."fileHeader".Keys | ForEach-Object {
        $headerCol = $_
        $foundValue = $hSetup."dbObject_mapDefault".($headerCol)
        # FIELD OVERRIDE BY USER.
        $containsOverride = $false
        if ($fieldOverrides.ContainsKey($headerCol)) {
            $foundValue = $fieldOverrides.($_)
            $containsOverride = $true
        }
        if ($foundValue) {
            # Perform a check to make sure a #N/A fields are not
            if ($foundValue.mapField -match "MANUALLY_REVIEW_AND_MAP_FIELD") {
                $hNewFieldMapping.($headerCol) = ($ifNoFieldIsFound + $tempCounter)
                $tempCounter++
            }
            else {
                $hNewFieldMapping.($headerCol) = $foundValue.mapField
                # must set override field here.
                if ($containsOverride -eq $true) {
                    $hNewFieldMapping.($headerCol) = $foundValue
                }
            }
        }
        else {
            $hNewFieldMapping.($headerCol) = ($ifNoFieldIsFound + $tempCounter)
            $tempCounter++
        }
    }

    # assign remap.
    $hSetup."fileHeaderRemap" = $hNewFieldMapping

    # access the correct worksheet
    $workSheet = $workbook.sheets.item("Mapped Header")

    # populate HEADER MAPPING
    # this applies header compact.
    $rowCounter = 3
    $hSetup."fileHeaderRemap".getEnumerator() | ForEach-Object {
        $key = $_.key
        $value = $-.value
        $fileHeaderUnAlteredHeaders = $hSetup."fileHeader".($key)
        # HEADER COMPACT
        $colCounters = 5..7
        foreach ($colCounter in $colCounters) {
    
            <# $colCounter is tcol$colCounters item #>
            switch ($colCounter) {
                5 { $rowValue = $fileHeaderUnAlteredHeaders }
                6 { $rowValue = $key }
                7 { $rowValue = $value }
            }
            $workSheet.cells.item($rowCounter, $colCounter) = $rowValue   
        }

        # HEADER EXTENDTED
        $valueSplits = $_.value -split "\|"
        foreach ($valueSplit in $valueSplits) {
            $colCounters = 1..3
            foreach ($colCounter in $colCounters) {    
                <# $colCounter is tcol$colCounters item #>
                switch ($colCounter) {
                    1 { $rowValue = $fileHeaderUnAlteredHeaders }
                    2 { $rowValue = $valueSplit }
                    3 { $rowValue = $key }
                }
                $workSheet.cells.item($rowCounter, $colCounter) = $rowValue
            }
            $rowCounter++
        }
    }

    # LEGACY check for unmapped headers
    $hitsFound = $hNewFieldMapping.GetEnumerator() | Where-Object { $_.value -match ($ifNoFieldIsFound + "\d+") }
    if ($hitsFound) {
        # starting recurse of file to pull example fields from all failed fields
        $counter = 0

        # field examples hash and create inital array.
        $numOfExample = 5
        $hFieldExamplesSkip = @{}
        $hFieldExamples = [ordered]@{}
        $hitsFound | ForEach-Object { $hFieldExamples.($_.Name) = @() }

        $import = Import-Csv $hSetup."filePath" -Encoding $hSetup."fileEncoding" -Delimiter $hSetup.fileDelimiterInput[0] -Header $hSetup["fileHeaderForImporting"]
        foreach ($elFields in $import) {
            <# $elFields is timp$import item #>
            # SKIP HEADER
            if ($counter -gt 0) {
                if ($hitsFound.Name.Count -ne $hFieldExamplesSkip.Count) {
                    # check all fields that couldn't be mapped
                    $hitsFound | ForEach-Object {
                        $fieldNotFoundName = $_.Name
                        $fieldValue = ($elFields.(($hSetup.fileDelimiterInput[1] + $fieldNotFoundName + $hSetup.fileDelimiterInput[1])) -replace $hSetup.fileDelimiterInput[1]).Trim() -replace "`t" # removes tab to make it easier putting value into excel.

                        if (!$hFieldExamplesSkip.ContainsKey($fieldNotFoundName)) {
                            if ($fieldValue.Length -gt 0) {
                                $currentExamples = $hFieldExamples.($fieldNotFoundName)
                                if (($currentExamples | Measure-Object).Count -lt $numOfExample) {
                                    $currentExamples += $fieldValue
                                    $hFieldExamples.($fieldNotFoundName) = $currentExamples
                                }
                                else {
                                    $hFieldExamplesSkip.($fieldNotFoundName) = ""
                                }
                            }
                        }
                    }
                }
                else {
                    break
                }
            }
            $counter++
        }
        # if all fields are empty, we don't want to stop the tool. We want it to proceed, however if 1 field has a single value the tool will stop and prompt the user.
        $checkForZeroHits = ($hFieldExamples.Values | Where-Object { $_ -match "." } | Measure-Object).Count
        if ($checkForZeroHits -eq 0) {
            $hitsFound.GetEnumerator() | ForEach-Object {
                $hSetup."fileHeaderRemap".($key) = "EMPTY_FIELD_IGNORE"
            }
            outputFieldExamplesToExcel $hFieldExamples, $numOfExample, $hitsFound "Unmapped Header" "continue"
        }
        else {
            # Output field examples for unmapped headers.
            outputFieldExamplesToExcel $hFieldExamples, $numOfExample, $hitsFound "Unmapped Header" "quit"
        }
    }

}

# FUNCTION - This will create a new file output.
Function fileRemapping {
    Function joinOutputline ($outLine) {
        $fieldSep = $hSetup.fileDelimiterOutput[0]
        $textSep = $hSetup.fileDelimiterOutput[1]
        $outLine -join ($textSep + $fieldSep + $textSep) -replace "^", $textSep -replace "$", ($textSep + $textSep)
    }

    $readEncode = [System.Text.Encoding]::UTF8
    $streamWriter = [System.IO.StreamWriter]::new(($hSetup.fileOutpath + '\' + $hSetup.fileName), $false, $readEncode)

    # store ignore examples for output.
    $numOfExample = 5
    # used for ignores values.
    $hFieldExamples = [ordered]@{}

    # without this, the memory consumption will go bonkers........ calling a hashtable directly is a big no no. Instead we're converting it to datatable.
    $fileHeaderRemap = New-Object System.Data.DataTable
    $fileHeaderRemap.Columns.Add("key") | Out-Null
    $fileHeaderRemap.Columns.Add("value") | Out-Null
    $fileHeaderRemap.PrimaryKey = $fileHeaderRemap.Columns[0]

    $hSetup.fileHeaderRemap.GetEnumerator() | ForEach-Object {
        $row = $fileHeaderRemap.NewRow()
        $row.key = $_.key
        $row.value = $_.value
        $fileHeaderRemap.Rows.Add($row)
    }

    # Initialize variables
    $rowCounter = -1
    $tempCounter = 0

    # Pre-allocate memory for hashtables
    $hTempRemap = @{}

    # Load CSV in batches
    $batchSize = 1000
    $csvLines = [System.IO.File]::ReadAllLines($hSetup."filePath")

    Write-Host ""
    Write-Host "START: PROCESSING LINES"

    for ($i = 0; $i -lt $csvLines.Length; $i += $batchSize) {
        $batch = $csvLines[$i..[Math]::Min($i + $batchSize - 1, $csvLines.Length - 1)]
        $batch | ForEach-Object {
            # Process each line in the batch
            $hTempRemap.Clear()
            if ($rowCounter -eq -1) {
                # Skip header row
            }
            else {
                # Process fields
                $properties = $_.Split($hSetup.fileDelimiterInput[0])
                for ($j = 0; $j -lt $properties.Length; $j++) {
                    $fieldName = compactHeaderFieldToRemoveExtraItem($properties[$j] -replace $hSetup.fileDelimiterInput[1])
                    $fieldValue = ($properties[$j] -replace $hSetup.fileDelimiterInput[1]).Trim()

                    # Remap fields and handle multi-values
                    $newFieldName = $fileHeaderRemap.rows.find($fieldName).value
                    $newFieldNameSplit = $newFieldName -split "\|"
                    foreach ($newField in $newFieldNameSplit) {
                        $found = $hSetup."dbObject_softwareFieldInfo".($newField)

                        # Handle date fields
                        if ($found.fieldType -eq "DATE" -and $fieldValue -match "00/00/0000|00000000") {
                            $fieldValue = ""
                        }

                        # Handle multi-value separators
                        $multiValueSep = " "
                        if ($found.multiValueSep -match ".") {
                            if ($found.multiValueSep -match 'datField') {
                                $multiValueSep = ' ' + $fieldName + ': '
                            }
                            else {
                                $multiValueSep = $found.multiValueSep
                            }
                        }

                        if ($hTempRemap.Contains($newField)) {
                            $newFieldValue = $hTempRemap[$newField] + $multiValueSep + $fieldValue
                            $hTempRemap[$newField] = $newFieldValue
                        }
                        else {
                            if (($multiValueSepDatFieldHandle -eq "false") -or (!$fieldValue)) {
                                $hTempRemap[$newField] = $fieldValue
                            }
                            else {
                                # First-time datField workflow
                                $hTempRemap[$newField] = ($multiValueSep -replace "^ ") + $fieldValue
                            }
                        }

                        # Log ignored fields
                        if ($newField -match "$IGNORE$") {
                            if (!$hFieldExamples.ContainsKey($fieldName)) {
                                $hFieldExamples[$fieldName] = @()
                            }
                            if ($fieldValue.Length -gt 0 -and ($hFieldExamples[$fieldName].Count -lt $numOfExamples)) {
                                $hFieldExamples[$fieldName] += $fieldValue
                            }
                        }
                    }
                }

                # Append extra fields
                foreach ($key in $hSetup."appendExtraFieldDATA".Keys) {
                    $hTempRemap[$key] = $hSetup."appendExtraFieldDATA"[$key]
                }

                # Write output
                if ($rowCounter -eq 0) {
                    $streamWriter.WriteLine((joinOutputline $hTempRemap.Keys))
                }
                $streamWriter.WriteLine((joinOutputline $hTempRemap.Values))

                # Logging in batches
                $tempCounter++
                if ($tempCounter -eq 50) {
                    Write-Host "Lines processed:" $rowCounter "/" $hSetup."fileTotalLines" ""
                    $tempCounter = 0
                }
            }
            $rowCounter++
        }

        Write-Host "Lines processed:" $rowCounter "/" $hSetup."fileTotalLines" ""
        Write-Host "FINISHED: PROCESSING LINES"
        $streamWriter.Close()

        # Output field examples for ignored fields.
        if ($hFieldExamples) {
            outputFieldExamplesToExcel $hFieldExamples $numOfExample "" "Ignore Field e.g." "continue"
        }

        # copy OPT file into folder
        Write-Host ""
        $aOPTFileTypes = @(".log", ".opt")
        foreach ($optFileType in $aOPTFileTypes) {
            <# $optFileType is topt$aOPTFileTypes item #>
            $optFile = $hSetup.("folderPath") + '\' + ($hSetup.("fileNameWithoutExt") + $optFileType)
            if ((Test-Path $optFile)) {
                Write-Host "COPYING: " $optFile
                Copy-Item $optFile $hSetup."fileOutpath" -Force
            }
        }

        # FUNCTION output load ready file info and compare line counts.
        outputLoadReadyFileInfo "Load Ready File INFO"

        if (!hSetup.("noPrompts")) {
            # FUNCTION Excel Object
            closeExcelObject

            if ($hSetup."fileAndTargetLineCountsMatch" -eq "true") {
                Write-Host "Header Norm Complete" -ForegroundColor Green
            }
            else {
                Write-Host "Header Norm Incomplete due to Incorrect Line Count" -ForegroundColor Red
            }

            Write-Host ""
            Write-Host "Output Path:"
            Write-Host $hSetup."fileOutpath"
            Write-Host ""
            Read-Host  "Press any key to exit"
        }
        else {

        }
    }


}
Function changeWindowSize {
    $pshost = Get-Host              # Get the PowerShell Host.
    $pswindow = $pshost.UI.RawUI    # Get the PowerShell Host's UI.

    $newsize = $pswindow. BufferSize    # Get the UI's current Buffer Size.
    $newsize.width = 128                # Set the new buffer's width to 150 columns.
    $pswindow.buffersize = $newsize     # Set the new Buffer Size as active.

    $newsize = $pswindow.windowsize     # Get the UI's current Window Size.
    $newsize.width = 128                # Set the new Window Width to 150 columns.
    $pswindow.windowsize = $newsize     # Set the new Window Size as active.
    $newsize = $pswindow.windowsize     # Get the UI's current Window Size.
    $newsize.height = 25                # Set the new Window Width to 150 columns.
}

Function setupUserArgs ($userArgs) {
    Function headerNormInfo {
        write-host "BASIC: " -foreground green -nonewline; "HEADER NORMALIZATION INFO"
        write-host ""
        write-host "EXAMPLE #1" -foreground cyan -nonewline; write-host " Basic Runtime for 99% of all users"
        write-host "Double click on headerNormTool.bat -- > SHIFT+RIGHT CLICK on the File -- > Paste into this Window -- > Hit Enter"
        write-host ""
        write-host "EXAMPLE #2" -foreground cyan -nonewline; write-host " Basic Runtime if a user doesn't want prompt (bulk running)"
        write-host "headerNormTool.bat -file \\UNC\DATFILE.dat -noPrompts"
        write-host ""
        write-host "EXAMPLE #3" -foreground cyan -nonewline; write-host " - Basic Runtime for Field Overide. Quickly remap fields outside of the DAT file."
        write-host "headerNormTool.bat -file \\UNC\DATFILE.dat -fieldOverride SOURCEPATH=SEC_SOURCE_CF,SOURCESHA=FILE HASHCODE SHA1, etc."
        write-host ""

        write-host ""
        $results = $hFlags. getEnumerator() | foreach-object {
            [PSCustomObject]@{
                'FLAG' = ("-" + $_.key)
                'INFO' = $_.value[1]
            }
        }
        ($results | format-table -auto | out-string).trim()
        write-host ""
    }

    # how many user arguments exist.
    $userArgsCount = ($userArgs | measure-object).count

    $hFlags = [ordered]@{
        "file"          = @("filePath", "Point to the DAT file to normalize.")
        "software"      = @("ediscoverySoftware", "Loading into Casepoint, NUIX, Law, etc .? Defaults to Casepoint.")
        "env"           = @("scriptEnv", "PROD, STAGE, DEV. Defaults to PROD")
        "fieldOverride" = @("fieldOverride", "If fields are unmapped and the template can't be updated. DATFIELD=SOFTWAREFIELD, DATFIELD=SOFTWAREFIELD")
        "bulkImport"    = @("bulkImport", "Point to a file that contains a collection of DAT files that need to be normalized.")
        "noPrompts"     = @("noPrompts", "Point to a file that contains a collection of DAT files that need to be normalized.")
    }
    # check and extract supported flags.
    $hSupportedFlags = [ordered]@{}
    $counter = 0

    # 99% of the time, people will just want to drag a DAT file, they don't need additional flags nor do they care. They just want to run and go.
    # To do this, we first check the user args, do flags exist? Does anything exist?
    # If flags exsit, we recurse and the approitae flags.
    # If flags don't exist, we only look for the \\ or F: |C: files, this way we work with only the Files for users. This will help streamline user input for the future.
    $hArgCheck = [ordered]@{
        "userArgsExist"     = "false"
        "anyThingPopulated" = "false"
    }

    foreach ($userArg in $userArgs) {
        if ($userArg) {
            $hArgCheck.("anyThingPopulated") = "true"
        }
        if ($userArg -match "^-") {
            $hArgCheck.("userArgsExist") = "true"
            break
        }
    }

    if (($hArgCheck.("userArgsExist") -eq "false") -and ($hArgCheck.("anyThingPopulated") -eq "false")) {
        # no data options selected.
        headerNormInfo $hFlags

        write-host "EXAMPLE #1 HAS BEEN SELECTED" -foreground magenta
        write-host ""
        [string]$filePath = read-host "Please Provide DAT File"
        $filePath = $filePath.replace('"', "").trim()
        Do {
            $tp = test-path $filePath
            if ($tp -match "false") {
                write-host ""
                write-host "!WARNING! DAT FILE NOT FOUND" -foreground red
                write-host ""
                $filePath = read-host "PROVIDE A DAT FILE"
            }
        }until ($tp -match "true")

        $hSupportedFlags. ("filePath") = $filePath
        $hSetup. ("userInputRun") = "true"

    }
    elseif ($hArgCheck.("userArgsExist") -eq "true") {
        # flags deteceted.
        foreach ($userArg in $userArgs) {
            foreach ($hFlag in $hFlags.getEnumerator()) {
                # populate with all found fields.
                if ($userArg -match ("-" + $hFlag.key)) {
                    $tempCounter = $counter + 1
                    $tempValue = ""
                    if ($tempCounter -le $userArgsCount) {
                        if ($userArg -match "-noPrompts") {
                            $tempValue = "true"
                        }
                        elseif ($userArg -match "-bulkImport") {
                            $tempValue = "true"
                        }
                        else {
                            $tempValue = $userArgs[$tempCounter].trim()
                        }
                    }
                    $hSupportedFlags. ($hFlag.value[0]) = $tempValue
                }
            }
            $counter++
        }
    }
    elseif ($hArgCheck.("anyThingPopulated") -eq "true") {
        foreach ($userArg in $userArgs) {
            if ($userArg -match "\\\\|:\\") {
                $hSupportedFlags.("filePath") = $userArg
            }
        }
    }
    # used to quickly populate some user arguments if they aren't provided.
    Function quickValidationForUserArgs ($found , $flag , $userArgToCheck , $updateWithIfBlank) {
        if ($flag -match $userArgToCheck) {    
            if (!$found) {
                $found = $updateWithIfBlank
            }
        }
        return $found
    }

    # run verification to make sure required values are hit.
    $hFlags. getEnumerator() | foreach-object {
        $flag = $_.value[0]
        $found = $hSupportedFlags.($flag)
        
        # if a header norm is used in the far future, removing this flag will allow people to easily update to use new ediscovery software.
        $found = quickValidationForUserArgs $found $flag "ediscoverySoftware" "casepoint"
        
        # if no script environment was specified, it's auto set to production.
        $found = quickValidationForUserArgs $found $flag "scriptEnv" "stage"
        
        # field override is necessary when a user needs to force certain field changes without needing to update a DAT file.
        $found = quickValidationForUserArgs $found $flag "fieldOverride" ""
        
        # bulk import parses a file to import all DAT files into a single if they all contain the same header.
        $found = quickValidationForUserArgs $found $flag "bulkImport" ""
        
        # add to main setup dictionary.
        $hSetup.($flag) = $found
    }
    
    # setup additional file output items.
    $hSetup.("folderPath") = $hSetup.("filePath") -replace "(^.+)\\.+$" , "`$1"
    $hSetup.("fileName") = $hSetup.("filePath") -replace "^.+\\(.+$)" , "`$1"
    $hSetup.("fileNameWithoutExt") = $hSetup.("fileName") -replace "(^.+)\..+$" , "`$1"
    $hSetup.("fileOutFolderName") = "NEW_CP DAT"
    $hSetup.("fileOutpath") = $hSetup.("folderPath") + '\' + $hSetup.("fileOutFolderName")
    $hSetup.("todayDate") = (get-date).toString("MM/dd/yyyy")

    # we need a proper file, in this case we know EMPTY is a no go, addition QC will be handled elsewhere.
    if ((!$hSetup.("filePath")) -and (!$hSetup.("bulkImport"))) {
        # no data options selected.
        headerNormInfo $hFlags

        write-host ""
        write-host "!ERROR!" -foreground red

        write-host "PROVIDE A PROPER EXAMPLE"
        write-host ""
        write-host "-filePath and -bulkImport are empty"
        exit
    }

    # verify filepath is proper before continuing.
    if (!(test-path $hSetup.("filePath"))) {
        # no data options selected.
        headerNormInfo $hFlags

        write-host ""
        write-host "!ERROR!" -foreground red

        write-host "PROVIDE A PROPER EXAMPLE"
        write-host ""
        write-host "-filePath does not exist:"
        write-host $hSetup. ("filePath")
        exit
    }
}

Function latestScriptVersion {
    # retrive the environment script that will be used. Prod uses \prod folder, Stage uses \stage, and so on...
    try {
        $folders = Get-ChildItem ($hSetup."scriptRootFolder" + '\' + $hSetup."scriptEnv") -Directory | Where-Object { $_.Name -match "^\d{5}$" } | Sort-Object Name -Descending | Select-Object -First 3 -ErrorAction Stop
    }
    catch {
        $except = $_.Exception.Message
    }   

    # verify folder has been found
    if (($except) -or ($folders.Count -eq 0)) {
        write-host ""
        write-host "!ERROR!" -foreground red
        write-host "Script not found, please attempt to re-run"
        write-host ""
        Read-Host  "Press any key to exit"
        exit
    }

    # retrived last versions, applying to setup table..
    foreach ($folder in $folders) {
        if (Test-Path ($folder.FullName + "\.lock")) {
            $latestFolder = $folder
            break
        }
    }

    $hSetup."scriptLatestVersionFolder" = $latestFolder.fullname
    $hSetup."scriptLatestVersion" = $latestFolder.name
    $hSetup."scriptConfig" = '_config.xml'
    $hSetup."scriptConfigPath" = ($latestFolder.fullname + '\' + $hSetup.scriptConfig)
    #$hSetup."mappingConfig" = ('mapping_' + $hSetup. "ediscoverySoftware" + '.xml')
    #$hSetup."mappingConfigPath" = ($latestFolder. fullname + '\' + $hSetup.mappingConfig)
    $hSetup."mappingInfoXLSXName" = "_mappingInfo.xlsx"
    $hSetup."mappingInfoXLSXNameOutName" = $hSetup.fileNameWithoutExt + "_mappingInfo.xlsx"
    $hSetup."dbPath" = $latestFolder.fullname + '\db.accdb'
}

# IMPORT MODULES For Scripts
Function scriptModuleImport {
    Get-ChildItem ($hSetup."scriptLatestVersionFolder" + '\scripts') -Filter "*.psm1" -Recurse -Exclude "tools" | ForEach-Object {
        $name = $_.Name -replace ".psm1"
        # import module
        try {
            Import-Module $_.FullName -ErrorAction Stop
        }
        catch {
            if ($error) {
                write-host "IMPORT MODULES"
                write-host "--------------------"
                write-host "$name`n"
                write-host "$_" -Foreground red
                write-host "`nScript exiting, Please fix module`n"
                Read-Host  "Press any key to exit"
                exit    
            }
        }
        finally {
            Write-Host "$name`n"
        }
    }
}

Clear-Host
# FUNCTION - changing window size
changeWindowSize

# define Variable Handoff
$hSetup = [ordered]@{}

# FUNCTION - validate user aguments
setupUserArgs $args

# machine script that the script is running from 
$hSetup."machineRunningFrom" = $env:COMPUTERNAME
# root location of the script
$hSetup."scriptRootFolder" = $PSScriptRoot

# FUNCTION - with user arguments set, we now the pull the latest script version for running
latestScriptVersion

# FUNCTION - retrive last version script modules.
scriptModuleImport

# FUNCTION - Open script configuration / workflow file.
$hSetup."scriptConfigXML" = openXMLFile $hSetup.scriptConfigPath

# FUNCTION - Open mapping DB and get the required tables. These will be used for all mappings going forward.
getDBObjects

# Create Output Directory
if (!(Test-Path $hSetup.fileOutpath)) {
    New-Item $hSetup.fileOutpath -ItemType Directory -Force | Out-Null
}

# copy mapping info xlsx.
Copy-Item ($hSetup.scriptLatestVersionFolder + '\' + $hSetup."mappingInfoXLSXName") ($hSetup.fileOutpath + '\' + $hSetup."mappingInfoXLSXNameOutName") -Force

# FUNCTION Excel Object
$excelObj = openExcelObject
$workBook = $excelObj.Workbooks.Open(($hSetup.fileOutpath + '\' + $hSetup."mappingInfoXLSXNameOutName"))

# LAUNCH Workflow Steps
$hSetup.ediscoverySoftware = "casepoint"
$xmlSearch = $hSetup.ediscoverySoftware
$hSetup."scriptConfigXML".selectNodes("//$xmlSearch").step | ForEach-Object {
    $_.module
}