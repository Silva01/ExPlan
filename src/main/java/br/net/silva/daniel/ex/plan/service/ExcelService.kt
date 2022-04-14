package br.net.silva.daniel.ex.plan.service

import br.net.silva.daniel.ex.plan.model.Excel
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import java.io.FileOutputStream

class ExcelService {

    private var linha : Int = 1

    private var workbook : HSSFWorkbook = HSSFWorkbook()

    @Throws(java.lang.Exception::class)
    fun novo(dados : List<List<Excel>>, destinoArquivo : String) {
        val sheet : HSSFSheet = workbook.createSheet()
        preencherTitulo(dados[0], sheet)
        dados.forEach { d -> preencher(d, sheet) }
        salvarArquivo(destinoArquivo)
    }

    private fun preencher(dados : List<Excel>, sheet : HSSFSheet) {
        dados.forEach { d -> preencherCelula(d.position(), d.value(), criarRow(sheet, incrementarLinha())) }
    }

    private fun preencherTitulo(dados : List<Excel>, sheet: HSSFSheet) {
        dados.forEach { d -> preencherCelula(d.position(), d.title(), criarRow(sheet, 0)) }
    }

    private fun incrementarLinha() : Int {
        return linha++
    }

    private fun preencherCelula(position : Int, value : Any, row : Row) {
        val cell : Cell = row.createCell(position)

        when(value) {
            is String -> cell.setCellValue(value)
            is Int -> cell.setCellValue(value.toDouble())
            is Double -> cell.setCellValue(value)
        }
    }

    private fun criarRow(sheet: HSSFSheet, rowPosicao : Int) : Row {
        return sheet.createRow(rowPosicao)
    }

    private fun salvarArquivo(destino : String) {
        val output = FileOutputStream(destino)
        workbook.write(output)
        output.close()
    }
}