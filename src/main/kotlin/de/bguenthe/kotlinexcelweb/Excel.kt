package de.bguenthe.kotlinexcelweb

import com.fasterxml.jackson.module.kotlin.jacksonObjectMapper
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.text.SimpleDateFormat

private val dateFormatter = SimpleDateFormat("dd/MM/yyyy")

fun main(args: Array<String>) {
//    val mapperlo = jacksonObjectMapper()
//    val numbers = listOf(1, 2, 3)
//    val strings = listOf("one", "two", "three")
//    val numStrList = numbers.zip(strings)
//
//    for (einer in numStrList) {
//        var json = mapperlo.writerWithDefaultPrettyPrinter().writeValueAsString(einer)
//        println(json)
//    }

    val filePath = "geschäftsvorfälle.xlsx"

    val file = File(filePath)

    val excelFileInputStream = FileInputStream(file)
    val readWorkbook = XSSFWorkbook(excelFileInputStream)
    val sheet = readWorkbook.getSheetAt(0)
    val header = readHeader(sheet)
    val data = readData(sheet)
    val mapper = jacksonObjectMapper()
    for (row in data) {
        val str = header.zip(row.value)
        val jsonStr = mapper.writerWithDefaultPrettyPrinter().writeValueAsString(str.toMap())
        println(jsonStr)
    }
}

fun readHeader(sheet: Sheet): ArrayList<String> {
    val header = ArrayList<String>()
    val row = sheet.getRow(0)
    for (cell in row)
        header.add(cell.toString())
    return header
}

fun readData(sheet: Sheet): MutableMap<Int, ArrayList<String>> {
    val map = mutableMapOf<Int, ArrayList<String>>() // zum füllen
    for (row in sheet) {
        val al = ArrayList<String>()
        if (row.rowNum == 0) // not for header
            continue
        for (cell in row) {
            al.add(cell.toString())
        }
        map[row.rowNum] = al
    }
    return map
}