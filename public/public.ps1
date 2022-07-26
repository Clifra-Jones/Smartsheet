<# using namespace System.Collections.Generic

enum cellProperties {
    conditionalFormat
    format
    formula
    hyperlink
    image
    linkInFromCell
    linksOutFromCell
    value
}


class Column {
    [string]$id
    [int]$version
    [int]$index
    [string]$title
    [string]$type
    [string]$description
    [string]$formula
    [bool]$validation
    [int]$width
    [bool]$hidden
    [psobject]$autoNumberFormat
    [psobject]$contactOptions
    [string]$format
    [bool]$locked
    [bool]$lockedForUser
    [string[]]$options
    [string]$symbol
    [string]$systemColumnType

    Column(){}

    Column(
        [string]$id,
        [int]$version,
        [int]$index,
        [string]$title,
        [string]$type,
        [string]$description,
        [string]$formula,
        [bool]$validation,
        [int]$width,
        [bool]$hidden,
        [psobject]$autoNumberFormat,
        [psobject]$contactOptions,
        [string]$format,
        [bool]$locked,
        [bool]$lockedForUser,
        [string[]]$options,
        [string]$symbol,
        [string]$systemColumnType
    ) 
    {
        $this.id = $id
        $this.version = $version
        $this.index = $index
        $this.title = $title
        $this.type = $type
        $this.description = $description
        $this.formula = $formula
        $this.validation = $validation
        $this.width = $width
        $this.hidden = $hidden
        $this.autoNumberFormat = $autoNumberFormat
        if (this.type -eq "CONTACT_LIST") {
            $this.$contactOptions = $contactOptions
        }
        $this.format = $format
        $this.locked = $locked
        $this.lockedForUser = $lockedForUser
        $this.options = $options
        $this.symbol = $symbol
        $this.systemColumnType = $systemColumnType
    }
    Column ([psobject]$Column) {
        $this.id = $Column.id
        $this.version = $Column.version
        $this.index = $Column.index
        $this.title = $Column.title
        $this.type = $Column.type
        $this.description = $Column.description
        $this.formula = $Column.formula
        $this.validation = $Column.validation
        $this.width = $Column.width
        $this.hidden = $Column.hidden
        $this.autoNumberFormat = $Column.autoNumberFormat
        if ($this.type -eq "CONTACT_LIST") {
            $this.contactOptions = $Column.contactOptions
        }
        $this.format = $Column.format
        $this.locked = $Column.locked
        $this.lockedForUser = $Column.lockedForUser
        $this.options = $Column.options
        $this.symbol = $Column.symbol
        $this.systemColumnType = $Column.systemColumnType        
    }
    Column (
        [int]$index,
        [string]$title,
        [string]$type
    )
    {
        $this.index = $index
        $this.title = $title
        $this.type = $type
    }
}

class Cell {
    [string]$columnid
    [string]$columnType
    [string]$conditionalFormat
    [string]$displayValue
    [string]$format
    [string]$formula
    [psobject]$hyperLink
    [psobject]$image
    [psobject]$linkInFromCell
    [psobject[]]$linksOutFromCell
    [psobject]$value

    Cell(){}

    Cell(
        [string]$columnId,
        [string]$columnType,
        [string]$conditionalFormat,
        [string]$displayValue,
        [string]$format,
        [string]$formula,
        [psobject]$hyperLink,
        [psobject]$image,
        [psobject]$linkInFromCell,
        [psobject[]]$linksOutFromCell,
        [psobject]$value
    )
    {
        $this.columnid = $columnId
        $this.columnType = $columnType
        $this.conditionalFormat = $conditionalFormat
        $this.displayValue = $displayValue
        $this.format = $format
        $this.formula = $formula
        $this.hyperLink = $hyperLink
        $this.image = $image
        $this.linkInFromCell = $linkInFromCell
        $this.linksOutFromCell = $linksOutFromCell
        $this.value = $value
    }

    Cell(
        [string]$columnId,
        [string]$value
    )
    {
        $this.columnid = $columnId
        $this.value = $this
    }

    Cell(
        [string]$columnId,
        [string]$value,
        [string]$format    
    )
    {
        $this.columnid = $columnId
        $this.value = $value
        $this.format = $format
    }
}

class Row {
    [string]$id
    [int]$rowNumber
    [bool]$expanded
    [datetime]$createdAt
    [datetime]$modifiedAt
    [string]$format
    [Cell[]]$Cells
    [bool]$locked

    Row() {}

    Row(
        [string]$id,
        [int]$rowNumber,
        [bool]$expanded,
        [datetime]$createdAt,
        [datetime]$modifiedAt,
        [Cell[]]$Cells,
        [string]$format,
        [bool]$locked
    )
    {
        $this.id = $id
        $this.rowNumber = $rowNumber
        $this.expanded = $expanded
        $this.createdAt = $createdAt
        $this.modifiedAt = $modifiedAt
        $this.Cells = $Cells
        $this.format = $format
        $this.locked = $locked
    }        

    Row(
        [bool]$expanded,
        [Cell[]]$Cells,
        [string]$format,
        [bool]$locked
    )
    {
        $this.expanded = $expanded
        $this.Cells = $Cells
        $this.format = $format
        $this.locked = $locked
    }        

    Row([psObject]$Row) {
        $this.id = $Row.id
        $this.rowNumber = $Row.rowNumber
        $this.createdAt = $Row.createdAt
        $this.modifiedAt = $Row.modifiedAt
        $this.expanded = $Row.expanded
        $this.Cells = $Row.Cells
        $this.format = $Row.format
        $this.locked = $Row.locked
    }        

    AddCell(
        [Cell]$Cell
    )
    {
        if (-not $this.Cells) {
            $this.Cells = @($Cell)
        } else {
            $this.Cells += $cell
        }
    }
} 

class Sheet {
    [string]$id
    [string]$Name
    [string]$version
    [int]$totalRowCount
    [string]$accessLevel
    [string[]]$effectiveAttachmentOptions
    [bool]$gantEnabled
    [bool]$dependenciesEnabled
    [bool]$resourceManagementEnabled
    [string]$resourceManagementType
    [bool]$cellImageUploadEnabled
    [psobject]$userSettings
    [string]$permalink
    [datetime]$createdAt
    [datetime]$modifiedAt
    [bool]$isMultiPicklistEnabled
    [System.Collections.Generic.List[psobject]]$Columns
    [System.Collections.Generic.List[psObject]]$Rows

    Sheet() {
        $this.Columns = [List[psObject]]::New()
        $this.Rows = [List[psObject]]::New()
        #$this.Columns = $_Columns
        #$this.Rows = $_Rows
    }
   Sheet($Sheet) {
        $this.Columns = [List[psObject]]::New()
        $this.Rows = [List[psObject]]::New()
        #$this.Columns = $_Columns
        #$this.Rows = $_Rows
        
        $this.id = $Sheet.id
        $this.Name = $Sheet.name     
        $this.version = $Sheet.version
        $this.totalRowCount = $Sheet.totalRowCount
        $this.accessLevel = $sheet.accessLevel
        $this.effectiveAttachmentOptions = $sheet.effectiveAttachmentOptions
        $this.gantEnabled = $sheet.ganttEnabled
        $this.dependenciesEnabled = $sheet.dependenciesEnabled
        foreach ($column in $Sheet.columns) {
            $this.Columns.Add([column]::New($column))
        }
        foreach ($row in $Sheet.rows) {
            $Rows.Add([Row]::New($row))
        }     
        
        $this.Columns = $_Columns
        $this.Rows = $_Rows
    }
    

    [psObject]ToPSObject() {
        $psRows = [List[psobject]]::New()

        foreach ($row in $this.Rows) {
            $psCells = @{}
            foreach ($cell in $row.Cells) {
                $columnName = ($this.Columns.Where({$_.Id -eq $cell.columnId})).title
                $psCells.Add($columnName, $cell.value)                
            }
            $psRows.Add($psCells)
        }
        
        return $psRows.ToArray()        
    }
  
}  
 #>