require 'java'

require_relative 'poi-bin-5.2.3//poi-excelant-5.2.3.jar'
require_relative 'poi-bin-5.2.3//poi-examples-5.2.3.jar'
require_relative 'poi-bin-5.2.3//poi-ooxml-full-5.2.3.jar'
require_relative 'poi-bin-5.2.3//poi-ooxml-5.2.3.jar'
require_relative 'poi-bin-5.2.3//poi-ooxml-lite-5.2.3.jar'
require_relative 'poi-bin-5.2.3//poi-5.2.3.jar'
require_relative 'poi-bin-5.2.3//lib/commons-io-2.11.0.jar'
require_relative 'poi-bin-5.2.3//lib/commons-collections4-4.4.jar'
require_relative 'poi-bin-5.2.3//lib/log4j-api-2.18.0.jar'
require_relative 'poi-bin-5.2.3//lib/commons-math3-3.6.1.jar'
require_relative 'poi-bin-5.2.3//lib/commons-codec-1.15.jar'
require_relative 'poi-bin-5.2.3//lib/SparseBitSet-1.2.jar'
require_relative 'poi-bin-5.2.3//poi-scratchpad-5.2.3.jar'
require_relative 'poi-bin-5.2.3//ooxml-lib/commons-logging-1.2.jar'
require_relative 'poi-bin-5.2.3//ooxml-lib/slf4j-api-1.7.36.jar'
require_relative 'poi-bin-5.2.3//ooxml-lib/jakarta.activation-2.0.1.jar'
require_relative 'poi-bin-5.2.3//ooxml-lib/jakarta.xml.bind-api-3.0.1.jar'
require_relative 'poi-bin-5.2.3//ooxml-lib/commons-compress-1.21.jar'
require_relative 'poi-bin-5.2.3//ooxml-lib/curvesapi-1.07.jar'
require_relative 'poi-bin-5.2.3//ooxml-lib/xmlbeans-5.1.1.jar'
require_relative 'poi-bin-5.2.3//poi-javadoc-5.2.3.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-xml-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-shared-resources-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-gvt-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/fontbox-2.0.26.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-codec-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-css-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/graphics2d-0.40.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-awt-util-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-svg-dom-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-svgrasterizer-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-i18n-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-parser-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-util-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/xmlgraphics-commons-2.6.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-constants-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-script-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/xml-apis-1.4.01.jar'
require_relative 'poi-bin-5.2.3//auxiliary/bcutil-jdk15on-1.70.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-svggen-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-ext-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-anim-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-transcoder-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-dom-1.14.jar'
require_relative 'poi-bin-5.2.3//auxiliary/bcpkix-jdk15on-1.70.jar'
require_relative 'poi-bin-5.2.3//auxiliary/pdfbox-2.0.26.jar'
require_relative 'poi-bin-5.2.3//auxiliary/xmlsec-3.0.0.jar'
require_relative 'poi-bin-5.2.3//auxiliary/bcprov-jdk15on-1.70.jar'
require_relative 'poi-bin-5.2.3//auxiliary/xml-apis-ext-1.3.04.jar'
require_relative 'poi-bin-5.2.3//auxiliary/batik-bridge-1.14.jar'

include Java

StWb = org.apache.poi.ss.usermodel.WorkbookFactory


# Open an existing workbook
file_path = 'sample.xlsx'
input_stream = Java::java.io.FileInputStream.new(file_path)
workbook = StWb.create(input_stream)

# Get the first sheet
sheet = workbook.getSheetAt(0)

cellB1 = sheet.getRow(0).getCell(1)
cellF1 = sheet.getRow(0).getCell(5)

evaluator = workbook.creationHelper.createFormulaEvaluator

cellB1.setCellValue(6)

# Evaluate the formula
cellValue = evaluator.evaluate(cellF1)

# Print the result
puts "SUM: #{cellValue.numberValue}"

cellB2 = sheet.getRow(1).getCell(1)
cellF2 = sheet.getRow(1).getCell(5)
cellValue = evaluator.evaluate(cellF2)
puts "Formula: #{cellValue.numberValue}"


cellB2 = sheet.getRow(1).getCell(1)
cellE2 = sheet.getRow(1).getCell(4)
cellF2 = sheet.getRow(1).getCell(5)
puts "E2: #{cellE2.numericCellValue}"
puts "Formula (No Update): #{cellF2.numericCellValue}"


cellB2.setCellValue(6)
evaluator = workbook.creationHelper.createFormulaEvaluator
cellValue = evaluator.evaluate(cellF2)
puts "Formula (Update): #{cellValue.numberValue}"


# Close the workbook
begin
  workbook.close
rescue Exception => e
  e.printStackTrace
end