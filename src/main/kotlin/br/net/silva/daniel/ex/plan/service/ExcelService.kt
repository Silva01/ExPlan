package br.net.silva.daniel.ex.plan.service

import br.net.silva.daniel.ex.plan.model.Excel
import br.net.silva.daniel.ex.plan.model.Read
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Workbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException
import java.nio.file.Files
import java.nio.file.Paths

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

    @Throws(java.lang.Exception::class)
    fun atualizar(dados : List<List<Excel>>, arquivoAtual: String, destinoArquivo : String) {
        val arquivo = FileInputStream(File(arquivoAtual))
        val work = HSSFWorkbook(arquivo)
        val sheet : HSSFSheet = work.getSheetAt(0)
        preencherTitulo(dados[0], sheet)
        dados.forEach { d -> preencher(d, sheet) }
        salvarArquivo(destinoArquivo, work)
    }

    fun mover(pathOrigem: String, pathDestino: String) = Files.move(Paths.get(pathOrigem), Paths.get(pathDestino))

    fun apagar(pathArquivo: String) = Files.delete(Paths.get(pathArquivo))

    @Throws(IOException::class)
    fun ler(destino : String) : List<List<Read>> {
        val arquivo = FileInputStream(File(destino))
        val work = HSSFWorkbook(arquivo)
        val sheet = work.getSheetAt(0)
        val lista = mutableListOf<Read>()

        val rowIterator = sheet.rowIterator()

        while (rowIterator.hasNext()) {
            val row = rowIterator.next()
            val cellIterator = row.cellIterator()

            while (cellIterator.hasNext()) {
                val cell = cellIterator.next()
                val read = Read(cell.rowIndex == 0, cell.stringCellValue, cell.columnIndex)
                lista.add(read)
            }
        }

        return mutableListOf(lista)
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

    private fun salvarArquivo(destino : String, workbook: Workbook) {
        val output = FileOutputStream(destino)
        workbook.write(output)
        output.close()
    }

}