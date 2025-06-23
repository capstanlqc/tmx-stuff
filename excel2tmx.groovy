/*:name = Excel2TMX :description =
 *                  This script converts XLS or XLSX file(s) located in <project>/script_input
 *                  to TMX file(s) and places them to <project>/tm/excel2tmx
 *                  The script extracts only the TUV with the same languages as set for the current project
 *                  Column headers in the spreadsheets 
 *
 * @author:     Kos Ivantsov
 * @date:       2025-06-23
 * @latest:     2025-06-23
 * @version:    1.0
 */

@Grab('org.apache.poi:poi:5.2.3')
@Grab('org.apache.poi:poi-ooxml:5.2.3')

import org.apache.poi.ss.usermodel.WorkbookFactory
import groovy.xml.MarkupBuilder
import groovy.xml.XmlUtil

// Variable alttype specifies which properties are used to mark alternative translations
// Supported values: 'id' and 'context' 
def alttype = "id"

// Default alttype to "id" if set to a not supported value
if (alttype != "context") {
    alttype = "id"
}
// Regex to define sheet names to be processed
def sheetPattern = ~/.*/

def props = project.projectProperties
if (!props) {
    console.println("No project open")
    return
}
def srcCode = props.sourceLanguage.languageCode
def tgtCode = props.targetLanguage.languageCode
def projectRoot = props.projectRoot

def inputDir = new File(projectRoot + File.separator + "script_input")
def outputDir = new File(props.getTMRoot() + File.separator + "excel2tmx")
if (!inputDir.exists()) {
    console.println("'script_input' folder not found")
    return
}
if (!outputDir.exists()) outputDir.mkdirs()

inputDir.listFiles(new FilenameFilter() {
    boolean accept(File dir, String name) {
        name.toLowerCase().endsWith('.xls') || name.toLowerCase().endsWith('.xlsx')
    }
}).each { file ->

    def data = []
    def workbook = WorkbookFactory.create(file)
    workbook.each { sheet ->
        if (!(sheet.sheetName ==~ sheetPattern)) return
        if (sheet.physicalNumberOfRows < 2) return
        def headerRow = sheet.getRow(1) // 2nd row (index 1)
        def headers = (0..<headerRow.lastCellNum).collect { i ->
            headerRow.getCell(i)?.toString()?.trim()
        }
        def segIdIdx = headers.indexOf('Segment ID')
        def srcIdx = headers.indexOf(srcCode)
        def tgtIdx = headers.indexOf(tgtCode)
        def altUniqIdx = headers.indexOf('Alt/Uniq')
        if ([segIdIdx, srcIdx, tgtIdx].any { it == -1 }) return

        def rows = []
        (0..<sheet.physicalNumberOfRows).each { rIdx ->
            def row = sheet.getRow(rIdx)
            if (row) rows << row
        }

        (2..<rows.size()).each { rIdx ->
            def row = rows[rIdx]
            def prevRow = rIdx > 1 ? rows[rIdx-1] : null
            def nextRow = rIdx < rows.size()-1 ? rows[rIdx+1] : null
            
            def targetText = row.getCell(tgtIdx)?.toString()
            def altUniq = altUniqIdx != -1 ? row.getCell(altUniqIdx)?.toString() : null
            def isForcedAlt = altUniq && altUniq.toLowerCase().contains('a')
            
            // Skip non-forced segments with empty targets
            if (!isForcedAlt && (targetText == null || targetText.trim().isEmpty())) {
                return
            }
            
            def sourceText = row.getCell(srcIdx)?.toString()
            def segmentId = row.getCell(segIdIdx)?.toString()
            def prevSource = prevRow?.getCell(srcIdx)?.toString() ?: ''
            def nextSource = nextRow?.getCell(srcIdx)?.toString() ?: ''
            
            // For 'context' or missing ID, collect context
            if (alttype == 'context' || !segmentId) {
                data << [
                    source_text: sourceText,
                    target_text: targetText,
                    prev_source: prevSource,
                    next_source: nextSource,
                    segment_id: null,
                    alt_uniq: altUniq
                ]
            } else {
                data << [
                    source_text: sourceText,
                    target_text: targetText,
                    segment_id: segmentId,
                    prev_source: null,
                    next_source: null,
                    alt_uniq: altUniq
                ]
            }
        }
    }
    workbook.close()

    // --- Categorization ---
    def grouped = data.groupBy { it.source_text }
    def defaultTranslations = []
    def alternativeTranslations = []

    grouped.each { source, items ->
        // Separate forced and non-forced items
        def forcedItems = items.findAll { it.alt_uniq?.toLowerCase()?.contains('a') }
        def nonForcedItems = items.findAll { !(it.alt_uniq?.toLowerCase()?.contains('a')) }
    
        // Add all forced items ONLY to alternativeTranslations
        alternativeTranslations.addAll(forcedItems)
    
        // Now process non-forced items for default/alternative logic
        if (nonForcedItems) {
            def targetCounts = nonForcedItems.countBy { it.target_text }
            def uniqueTargets = targetCounts.keySet().toList()
    
            if (uniqueTargets.size() == 1) {
                defaultTranslations << [source_text: source, target_text: uniqueTargets[0]]
            } else {
                def maxCount = targetCounts.values().max()
                def maxTargets = targetCounts.findAll { k, v -> v == maxCount }.keySet()
                // Pick the first encountered as default
                def defaultTarget = null
                for (item in nonForcedItems) {
                    if (maxTargets.contains(item.target_text)) {
                        defaultTarget = item.target_text
                        break
                    }
                }
                if (defaultTarget != null) {
                    defaultTranslations << [source_text: source, target_text: defaultTarget]
                }
                // Add all non-forced items for other targets to alternatives
                uniqueTargets.each { variantTarget ->
                    if (variantTarget == defaultTarget) return
                    nonForcedItems.findAll { it.target_text == variantTarget }.each {
                        alternativeTranslations << it
                    }
                }
            }
        }
        // If there are only forced items (no non-forced), do not add to default
    }

    
    // Add forced alternatives
    data.each { item ->
        if (item.alt_uniq && item.alt_uniq.toLowerCase().contains('a')) {
            alternativeTranslations << item
        }
    }
    
    // Deduplicate alternative translations
    def seen = [] as Set
    def filteredAlternatives = []
    alternativeTranslations.each { item ->
        def key
        if (alttype == 'context' || !item.segment_id) {
            key = [item.source_text, item.target_text, item.prev_source, item.next_source]
        } else {
            key = [item.source_text, item.target_text, item.segment_id]
        }
        
        if (!seen.contains(key)) {
            seen << key
            filteredAlternatives << item
        }
    }
    alternativeTranslations = filteredAlternatives

    // --- TMX Generation ---
    def writer = new StringWriter()
    def xml = new MarkupBuilder(writer)
    xml.mkp.xmlDeclaration(version: "1.0", encoding: "UTF-8")
    xml.tmx(version: '1.4') {
        header(
            creationtool: 'Excel2TMX',
            creationtoolversion: '1.0',
            segtype: 'sentence',
            adminlang: 'en-us',
            srclang: srcCode,
            datatype: 'PlainText'
        )
        body {
            defaultTranslations.each { item ->
                tu {
                    tuv('xml:lang': srcCode) {
                        seg(item.source_text ?: '')
                    }
                    tuv('xml:lang': tgtCode) {
                        seg(item.target_text ?: '')
                    }
                }
            }
            mkp.comment('Alternative translations')
            alternativeTranslations.each { item ->
                tu {
                    prop(type: 'file', file.name)
                    if (alttype == 'id' && item.segment_id) {
                        prop(type: 'id', item.segment_id)
                    } else {
                        prop(type: 'prev', item.prev_source ?: '')
                        prop(type: 'next', item.next_source ?: '')
                    }
                    tuv('xml:lang': srcCode) {
                        seg(item.source_text ?: '')
                    }
                    tuv('xml:lang': tgtCode) {
                        // Forced alternatives get empty seg if target is empty
                        if (item.alt_uniq?.toLowerCase()?.contains('a') && 
                           (item.target_text == null || item.target_text.trim().isEmpty())) {
                            seg('')
                        } else {
                            seg(item.target_text ?: '')
                        }
                    }
                }
            }
        }
    }

    // --- Pretty Print and Write ---
    def prettyXml = XmlUtil.serialize(writer.toString())
    // Remove empty lines
    prettyXml = prettyXml.replaceAll(/(?m)^[ \t\r\f]*\n/, '')
    def baseName = file.name.replaceFirst(/(?i)\.xlsx?$/, '')
    def outputPath = new File(outputDir, baseName + ".tmx")
    if ((defaultTranslations.size() == 0) && (alternativeTranslations.size() == 0)) {
        console.println("Nothing was extracted from the Excel file(s), TMX is not written")
    } else {
        outputPath.text = prettyXml
        console.println "TMX file created at: ${outputPath}"
        console.println("Default translations: ${defaultTranslations.size()}\nAlternative translations: ${alternativeTranslations.size()}")
    }
}
console.println("The script Excel2TMX finished")
