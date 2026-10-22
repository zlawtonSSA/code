#import csv files
$inputDir  = "C:\_ParentSquare\Unprocessed"
$outputDir = "C:\_ParentSquare\Processed"
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

$parents  = Import-Csv (Get-ChildItem "$inputDir\*parents*.csv"  | Select-Object -First 1 -ExpandProperty FullName)
$rosters  = Import-Csv (Get-ChildItem "$inputDir\*rosters*.csv"  | Select-Object -First 1 -ExpandProperty FullName)
$schools  = Import-Csv (Get-ChildItem "$inputDir\*schools*.csv"  | Select-Object -First 1 -ExpandProperty FullName)
$sections = Import-Csv (Get-ChildItem "$inputDir\*sections*.csv" | Select-Object -First 1 -ExpandProperty FullName)
$staff    = Import-Csv (Get-ChildItem "$inputDir\*staff*.csv"    | Select-Object -First 1 -ExpandProperty FullName)
$students = Import-Csv (Get-ChildItem "$inputDir\*students*.csv" | Select-Object -First 1 -ExpandProperty FullName)
$terms    = Import-Csv (Get-ChildItem "$inputDir\*terms*.csv"    | Select-Object -First 1 -ExpandProperty FullName)

Write-Host "Loaded:"
Write-Host "  parents:  $($parents.Count) records"
Write-Host "  rosters:  $($rosters.Count) records"
Write-Host "  schools:  $($schools.Count) records"
Write-Host "  sections: $($sections.Count) records"
Write-Host "  staff:    $($staff.Count) records"
Write-Host "  students: $($students.Count) records"
Write-Host "  terms:    $($terms.Count) records"

#region Schools Transform
$schoolDetails = @{
    '1' = @{ zip = '15238'; address = '423 Fox Chapel Rd'; city = 'Pittsburgh'; state = 'PA'; principal_first = 'Andrew'; principal_last = 'Prince'; principal_email = 'aprince@shadysideacademy.org'; phone = '412-968-3000' }
    '2' = @{ zip = '15238'; address = '423 Fox Chapel Rd'; city = 'Pittsburgh'; state = 'PA'; principal_first = 'Michelle'; principal_last = 'Merson'; principal_email = 'mmerson@shadysideacademy.org'; phone = '412-968-3100' }
    '3' = @{ zip = '15221'; address = '400 S Braddock Ave'; city = 'Pittsburgh'; state = 'PA'; principal_first = 'Dan'; principal_last = 'Stern'; principal_email = 'dstern@shadysideacademy.org'; phone = '412-968-3179' }
    '4' = @{ zip = '15238'; address = '400 Christ Curch Ln'; city = 'Pittsburgh'; state = 'PA'; principal_first = 'Beth'; principal_last = 'Grover'; principal_email = 'bgrover@shadysideacademy.org'; phone = '412-936-8644' }
}

$schoolsOut = $schools | ForEach-Object {
    $d = $schoolDetails[$_.school_id]
    [PSCustomObject]@{
        school_id             = $_.school_id
        school_name           = $_.school_name
        school_zip            = $d.zip
        school_address        = $d.address
        school_city           = $d.city
        school_state          = $d.state
        principal_first_name  = $d.principal_first
        principal_last_name   = $d.principal_last
        principal_email       = $d.principal_email
        school_phone          = $d.phone
    }
}


#endregion

#region Students Transform
$gradeMap = @{
    'JPK' = '-2'; 'SPK' = '-1'; 'KG'  = '0'
    '1st' = '1';  '2nd' = '2';  '3rd' = '3';  '4th' = '4'
    '5th' = '5';  '6th' = '6';  '7th' = '7';  '8th' = '8'
    '9th' = '9';  '10th'= '10'; '11th'= '11'; '12th'= '12'
}

$studentsOut = $students | ForEach-Object {
    [PSCustomObject]@{
        school_id       = $_.school_id
        student_id      = $_.student_id
        first_name      = $_.first_name
        last_name       = $_.last_name
        grade_level     = $gradeMap[$_.grade_level]
        # only include student emails that contain @shadysideacademy.org'
        student_email   = if ($_.student_email -like "*@shadysideacademy.org") { $_.student_email } else { $null }
        cellphone       = ($_.cellphone -replace '\D', '')
    }
}
#endregion

#region Parents Transform
$studentSchoolMap = @{}
$students | ForEach-Object {
    $studentSchoolMap[$_.student_id] = $_.school_id
}

$parentsOut = $parents | ForEach-Object {
    [PSCustomObject]@{
        school_id   = $studentSchoolMap[$_.student_id]
        student_id  = $_.student_id
        parent_id   = $_.parent_id
        first_name  = $_.first_name
        last_name   = $_.last_name
        email       = $_.email
        mobile      = ($_.mobile -replace '\D', '')
    }
}
#endregion

#region Terms Transform
$portionMap = @{
    'Semesters'           = 'Semester'
    'MS Trimesters'       = 'Trimester'
    'SS Semesters'        = 'Semester'
    'MS Semesters'        = 'Semester'
    'Athletic Trimesters' = 'Trimester'
    'Maymester'           = ''
    'Summer'              = ''
    '<None Specified>'    = ''
}

# Build term_id -> set of school_ids from sections (skip blank school_ids)
$termSchoolMap = @{}
$sections | Where-Object { $_.school_id -ne '' } | ForEach-Object {
    if (-not $termSchoolMap.ContainsKey($_.term_id)) {
        $termSchoolMap[$_.term_id] = [System.Collections.Generic.HashSet[string]]::new()
    }
    [void]$termSchoolMap[$_.term_id].Add($_.school_id)
}

$termLookup = @{}
$terms | ForEach-Object { $termLookup[$_.term_id] = $_ }

$termsOut = foreach ($termId in $termSchoolMap.Keys) {
    $t = $termLookup[$termId]
    if (-not $t) { continue }

    # Convert MM/DD/YY to MM/DD/YYYY
    $sp = $t.'Grading Period Begin Date' -split '/'
    $ep = $t.'Grading Period End Date'   -split '/'
    $startDate = "$($sp[0])/$($sp[1])/20$($sp[2])"
    $endDate   = "$($ep[0])/$($ep[1])/20$($ep[2])"

    # Derive school_year from start date
    $startMonth = [int]$sp[0]
    $startYear  = [int]"20$($sp[2])"
    if ($startMonth -ge 7) {
        $schoolYear = "$startYear-$(($startYear + 1).ToString().Substring(2))"
    } else {
        $schoolYear = "$($startYear - 1)-$(($startYear).ToString().Substring(2))"
    }

    foreach ($schoolId in $termSchoolMap[$termId]) {
        [PSCustomObject]@{
            school_id   = $schoolId
            term_id     = $t.term_id
            term_name   = $t.Description
            portion     = if ($portionMap.ContainsKey($t.'Grading Period Group')) { $portionMap[$t.'Grading Period Group'] } else { '' }
            start_date  = $startDate
            end_date    = $endDate
            school_year = $schoolYear
        }
    }
}
#endregion

#region Sections Transform
$sectionsOut = $sections | ForEach-Object {
    [PSCustomObject]@{
        school_id      = $_.school_id
        section_id     = $_.section_id
        term_id        = $_.term_id
        staff_id       = $_.staff_id
        course_name    = $_.course_name
        section_number = $_.section_number
        is_primary     = 1
    }
}
#endregion

#region Staff Transform
$staffOut = $staff | ForEach-Object {
    [PSCustomObject]@{
        school_id  = $_.school_id
        staff_id   = $_.staff_id
        title      = $_.title
        first_name = $_.first_name
        last_name  = $_.last_name
        email      = $_.email
        mobile     = ($_.mobile -replace '\D', '')
    }
}
#endregion

#region Rosters Transform
$rostersOut = $rosters | ForEach-Object {
    [PSCustomObject]@{
        school_id  = $_.school_id
        section_id = $_.section_id
        student_id = $_.student_id
    }
}
#endregion

#region Export
$schoolsOut  | Export-Csv "$outputDir\schools.csv"  -NoTypeInformation
$studentsOut | Export-Csv "$outputDir\students.csv" -NoTypeInformation
$parentsOut  | Export-Csv "$outputDir\parents.csv"  -NoTypeInformation
$termsOut    | Export-Csv "$outputDir\terms.csv"    -NoTypeInformation
$sectionsOut | Export-Csv "$outputDir\sections.csv" -NoTypeInformation
$staffOut    | Export-Csv "$outputDir\staff.csv"    -NoTypeInformation
$rostersOut  | Export-Csv "$outputDir\rosters.csv"  -NoTypeInformation

Write-Host "Exported to $outputDir"
#endregion