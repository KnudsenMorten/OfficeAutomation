#------------------------------------------------------------------------------------------
# IT Operating Model | Generate Powerpoint slides based on Excel speadsheet
#------------------------------------------------------------------------------------------

# Variables
    $CsvPath   = $ENV:OneDrive + "\Desktop\TEMP\Build_PPT_from_CSV\" + "IT Operating Services and Technologies, Tasks, KPIs.csv"
    $ExcelPath = $ENV:OneDrive + "\Desktop\TEMP\Build_PPT_from_CSV\" + "IT Operating Services and Technologies, Tasks, KPIs.xlsx"

# Import the CSV file
    $Data = Import-Csv -Path $csvPath -Delimiter ";"
    # $Data = import-excel $ExcelPath -WorksheetName 

#------------------------------------------------------------------------------------------
# KPI Overview
#------------------------------------------------------------------------------------------

# Variables
    $Presentationpath = $ENV:OneDrive + "\Desktop\TEMP\Build_PPT_from_CSV\" + "IT_Operating_Services_KPIs.pptx"

# Build PPT
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Add()

# Slide
    $Operation_Category  = ""
    $Service_Deliverable = ""
    $SlideCount = 0

    ForEach ($Row in $Data)
        {
               If ($Row.Service_Deliverable -ne $Service_Deliverable)
                {
                    $CountRows = (($Data | Where-Object { ($_.Service_Deliverable -eq $Row.Service_Deliverable) }) | Measure-Object).count
                    $SlideCount++
                    $Slide = $presentation.Slides.addSlide($presentation.Slides.Count + 1, $presentation.SlideMaster.CustomLayouts(2))
                    $Slide.Shapes(1).TextFrame.TextRange.Text = "$($Row.Service_Deliverable)"
                    $Slide.Shapes(2).TextFrame.TextRange.Text = "Category: $($Row.Operation_Main_Category) | $($Row.Operation_Sub_Category) - Environment: $($Row.Environment)"

                    # Table
                    $Table = $Slide.Shapes.AddTable($CountRows,5,45,210)
                    $Table.table.Columns.Item(1).Width = 150
                    $Table.table.Columns.Item(2).Width = 290
                    $Table.table.Columns.Item(3).Width = 290
                    $Table.table.Columns.Item(4).Width = 75
                    $Table.table.Columns.Item(5).Width = 75

                    # Build header
                    $FontSize = 8

	                $Table.table.Cell(1,1).shape.TextFrame.TextRange.Font.size = $FontSize
                    $Table.table.Cell(1,1).Shape.TextFrame.TextRange.Text = "KPI Name"

	                $Table.table.Cell(1,2).shape.TextFrame.TextRange.Font.size = $FontSize
                    $Table.table.Cell(1,2).Shape.TextFrame.TextRange.Text = "KPI Target Description"

	                $Table.table.Cell(1,3).shape.TextFrame.TextRange.Font.size = $FontSize
                    $Table.table.Cell(1,3).shape.TextFrame.TextRange.Text = "KPI Target"

	                $Table.table.Cell(1,4).shape.TextFrame.TextRange.Font.size = $FontSize
                    $Table.table.Cell(1,4).Shape.TextFrame.TextRange.Text = "KPI Category"

	                $Table.table.Cell(1,5).shape.TextFrame.TextRange.Font.size = $FontSize
                    $Table.table.Cell(1,5).Shape.TextFrame.TextRange.Text = "KPI Period"
                    $TableRows = 1

                    $temp = $Table.table.Rows.Add()
                }

                #Add row
                $TableRows++
                $Table.table.Cell($TableRows,1).Shape.TextFrame.TextRange.Font.size = $FontSize
                $Table.table.Cell($TableRows,1).shape.TextFrame.TextRange.Text = $Row.KPI_Name

                $Table.table.Cell($TableRows,2).Shape.TextFrame.TextRange.Font.size = $FontSize
                $Table.table.Cell($TableRows,2).shape.TextFrame.TextRange.Text = $Row.KPI_Target_Description

                $Table.table.Cell($TableRows,3).Shape.TextFrame.TextRange.Font.size = $FontSize
                $Table.table.Cell($TableRows,3).shape.TextFrame.TextRange.Text = $Row.KPI_Target

                $Table.table.Cell($TableRows,4).Shape.TextFrame.TextRange.Font.size = $FontSize
                $Table.table.Cell($TableRows,4).shape.TextFrame.TextRange.Text = $Row.KPI_Category_WHY

                $Table.table.Cell($TableRows,5).Shape.TextFrame.TextRange.Font.size = $FontSize
                $Table.table.Cell($TableRows,5).shape.TextFrame.TextRange.Text = $Row.KPI_Measurement_Period

                # Auto-adjust the row height
                $Table.table.Rows.Item($TableRows).Height = -1

                $Service_Deliverable = $Row.Service_Deliverable
        }

# Save the presentation
    $presentation.SaveAs($Presentationpath)

    Start-Sleep -s 5
    $presentation.Close()

    write-host "KPIs - Number of Slides: $($SlideCount)"
    Start-Sleep -s 3

#------------------------------------------------------------------------------------------
# KPI Implementation
#------------------------------------------------------------------------------------------

# Variables
    $Presentationpath = $ENV:OneDrive + "\Desktop\TEMP\Build_PPT_from_CSV\" + "IT_Operating_Services_KPI_Implementations.pptx"

# Build PPT
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Add()

# Data scope
    $DataScope = Sort-Object -InputObject $Data -Property Priority_Phase_Implementation

# Slide
    $Priority_Phase_Implementation  = ""
    $SlideCount = 0

    ForEach ($Row in $DataScope)
        {
               If ($Row.Priority_Phase_Implementation -ne $Priority_Phase_Implementation)
                {

                    If ($Row.Priority_Phase_Implementation)
                        {
                            $CountRows = (($DataScope | Where-Object { ($_.Priority_Phase_Implementation -eq $Row.Priority_Phase_Implementation) }) | Measure-Object).count
                        }
                    Else
                        {
                            $CountRows = 15
                        }
                    $SlideCount++
                    $Slide = $presentation.Slides.addSlide($presentation.Slides.Count + 1, $presentation.SlideMaster.CustomLayouts(2))
                    $Slide.Shapes(1).TextFrame.TextRange.Text = "$($Row.Priority_Phase_Implementation)"

                    # Table
                    $Table = $Slide.Shapes.AddTable($CountRows,4,45,210)
                    $Table.table.Columns.Item(1).Width = 150
                    $Table.table.Columns.Item(2).Width = 75
                    $Table.table.Columns.Item(3).Width = 290
                    $Table.table.Columns.Item(4).Width = 290

                    # Build header
                    $FontSize = 8

	                $Table.table.Cell(1,1).shape.TextFrame.TextRange.Font.size = $FontSize
                    $Table.table.Cell(1,1).Shape.TextFrame.TextRange.Text = "KPI Name"

	                $Table.table.Cell(1,2).shape.TextFrame.TextRange.Font.size = $FontSize
                    $Table.table.Cell(1,2).Shape.TextFrame.TextRange.Text = "KPI Impl. Complexity"

	                $Table.table.Cell(1,3).shape.TextFrame.TextRange.Font.size = $FontSize
                    $Table.table.Cell(1,3).shape.TextFrame.TextRange.Text = "KPI Implementation"

	                $Table.table.Cell(1,4).shape.TextFrame.TextRange.Font.size = $FontSize
                    $Table.table.Cell(1,4).Shape.TextFrame.TextRange.Text = "KPI Blockers"

                    $TableRows = 1

                    $temp = $Table.table.Rows.Add()
                }

                #Add row
                $TableRows++
                $Table.table.Cell($TableRows,1).Shape.TextFrame.TextRange.Font.size = $FontSize
                $Table.table.Cell($TableRows,1).shape.TextFrame.TextRange.Text = $Row.KPI_Name

                $Table.table.Cell($TableRows,2).Shape.TextFrame.TextRange.Font.size = $FontSize
                $Table.table.Cell($TableRows,2).shape.TextFrame.TextRange.Text = $Row.KPI_Implementation_Complexity

                $Table.table.Cell($TableRows,3).Shape.TextFrame.TextRange.Font.size = $FontSize
                $Table.table.Cell($TableRows,3).shape.TextFrame.TextRange.Text = $Row.KPI_Implementation

                $Table.table.Cell($TableRows,4).Shape.TextFrame.TextRange.Font.size = $FontSize
                $Table.table.Cell($TableRows,4).shape.TextFrame.TextRange.Text = $Row.KPI_Blockers

                # Auto-adjust the row height
                $Table.table.Rows.Item($TableRows).Height = -1

                $Priority_Phase_Implementation = $Row.Priority_Phase_Implementation
        }

# Save the presentation
    $presentation.SaveAs($Presentationpath)

    Start-Sleep -s 5
    $presentation.Close()

    write-host "KPI Implementation - Number of Slides: $($SlideCount)"
    Start-Sleep -s 3

#------------------------------------------------------------------------------------------
# Operational Tasks
#------------------------------------------------------------------------------------------

# Variables
    $Presentationpath = $ENV:OneDrive + "\Desktop\TEMP\Build_PPT_from_CSV\" + "IT_Operating_Services_Operational_Tasks.pptx"

# Build PPT
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Add()

# Slide
    $Operation_Category  = ""
    $Service_Deliverable = ""
    $SlideCount = 0

    ForEach ($Row in $Data)
        {
               If ($Row.Service_Deliverable -ne $Service_Deliverable)
                {
                    $CountRows = (($Data | Where-Object { ($_.Service_Deliverable -eq $Row.Service_Deliverable) }) | Measure-Object).count

                    $SlideCount++
                    $Slide = $presentation.Slides.addSlide($presentation.Slides.Count + 1, $presentation.SlideMaster.CustomLayouts(2))
                    $Slide.Shapes(1).TextFrame.TextRange.Text = "$($Row.Service_Deliverable)"
                    $Slide.Shapes(2).TextFrame.TextRange.Text = "Category: $($Row.Operation_Main_Category) | $($Row.Operation_Sub_Category) - Environment: $($Row.Environment)"

                    # Table
                    $Table = $Slide.Shapes.AddTable($CountRows,4,45,210)
                    $Table.table.Columns.Item(1).Width = 150
                    $Table.table.Columns.Item(2).Width = 320
                    $Table.table.Columns.Item(3).Width = 75
                    $Table.table.Columns.Item(4).Width = 320

                    # Build header
                    $FontSize = 8

	                $Table.table.Cell(1,1).shape.TextFrame.TextRange.Font.size = $FontSize
                    $Table.table.Cell(1,1).Shape.TextFrame.TextRange.Text = "KPI Name"

	                $Table.table.Cell(1,2).shape.TextFrame.TextRange.Font.size = $FontSize
                    $Table.table.Cell(1,2).Shape.TextFrame.TextRange.Text = "Operational Task"

	                $Table.table.Cell(1,3).shape.TextFrame.TextRange.Font.size = $FontSize
                    $Table.table.Cell(1,3).Shape.TextFrame.TextRange.Text = "Frequency"

	                $Table.table.Cell(1,4).shape.TextFrame.TextRange.Font.size = $FontSize
                    $Table.table.Cell(1,4).shape.TextFrame.TextRange.Text = "PreReq Operational Task"

                    $TableRows = 1

                    $temp = $Table.table.Rows.Add()
                }

                #Add row
                $TableRows++

                $Table.table.Cell($TableRows,1).Shape.TextFrame.TextRange.Font.size = $FontSize
                $Table.table.Cell($TableRows,1).shape.TextFrame.TextRange.Text = $Row.KPI_Name

                $Table.table.Cell($TableRows,2).Shape.TextFrame.TextRange.Font.size = $FontSize
                $Table.table.Cell($TableRows,2).shape.TextFrame.TextRange.Text = $Row.Operational_Task

                $Table.table.Cell($TableRows,3).Shape.TextFrame.TextRange.Font.size = $FontSize
                $Table.table.Cell($TableRows,3).shape.TextFrame.TextRange.Text = $Row.OprTask_Frequency

                $Table.table.Cell($TableRows,4).Shape.TextFrame.TextRange.Font.size = $FontSize
                $Table.table.Cell($TableRows,4).shape.TextFrame.TextRange.Text = $Row.PreReq_Operational_Task

                # Auto-adjust the row height
                $Table.table.Rows.Item($TableRows).Height = -1

                $Service_Deliverable = $Row.Service_Deliverable
        }

# Save the presentation
    $presentation.SaveAs($Presentationpath)

    Start-Sleep -s 5
    $presentation.Close()
    $powerPoint.Quit()

    write-host "Operational Tasks - Number of Slides: $($SlideCount)"


# Close PowerPoint
