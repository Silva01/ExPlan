package br.net.silva.daniel.ex.plan.test.service

import br.net.silva.daniel.ex.plan.model.Excel
import br.net.silva.daniel.ex.plan.service.ExcelService
import br.net.silva.daniel.ex.plan.test.model.DadoTeste
import org.junit.jupiter.api.Assertions
import org.junit.jupiter.api.Test
import java.io.File

class ExcelServiceTest {

    @Test
    fun deveCriarUmaPlanilha() {
        val dado: DadoTeste = DadoTeste("teste", "Opa", 0)
        val lista : List<Excel> = mutableListOf(dado)
        val dadosPlanilha : List<List<Excel>> = mutableListOf(lista)

        val service : ExcelService = ExcelService()
        service.novo(dadosPlanilha, "teste.xls")

        val planilha = File("teste.xls")

        Assertions.assertTrue(planilha.exists())
    }
}