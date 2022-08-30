package br.net.silva.daniel.ex.plan.test.service

import br.net.silva.daniel.ex.plan.model.Excel
import br.net.silva.daniel.ex.plan.service.ExcelService
import br.net.silva.daniel.ex.plan.test.model.DadoTeste
import org.junit.jupiter.api.Assertions
import org.junit.jupiter.api.MethodOrderer
import org.junit.jupiter.api.Order
import org.junit.jupiter.api.Test
import org.junit.jupiter.api.TestMethodOrder
import java.io.File
import java.util.Objects

@TestMethodOrder(MethodOrderer.OrderAnnotation::class)
class ExcelServiceTest {

    @Test
    @Order(1)
    fun deveCriarUmaPlanilha() {
        val dado: DadoTeste = DadoTeste("teste", "Opa", 0)
        val lista : List<Excel> = mutableListOf(dado)
        val dadosPlanilha : List<List<Excel>> = mutableListOf(lista)

        val service : ExcelService = ExcelService()
        service.novo(dadosPlanilha, "teste.xls")

        val planilha = File("teste.xls")

        Assertions.assertTrue(planilha.exists())
    }

    @Test
    @Order(2)
    fun deveLerOsDadosDaPlanilha() {
        val service = ExcelService()
        val dado = service.ler("teste.xls")

        Assertions.assertTrue(Objects.nonNull(dado))
        Assertions.assertEquals(1, dado.size)
        Assertions.assertTrue(dado[0][0].title)
        Assertions.assertEquals("teste", dado[0][0].value)
        Assertions.assertEquals(0, dado[0][0].position)
        Assertions.assertEquals("Opa", dado[0][1].value)

    }

    @Test
    @Order(3)
    fun deveAtualizarUmaPlanilha() {
        val dado: DadoTeste = DadoTeste("teste", "Isso e um teste unitario", 0)
        val lista : List<Excel> = mutableListOf(dado)
        val dadosPlanilha : List<List<Excel>> = mutableListOf(lista)

        val service : ExcelService = ExcelService()
        service.atualizar(dadosPlanilha, "teste.xls", "update.xls")

        val planilha = File("update.xls")

        Assertions.assertTrue(planilha.exists())

        val dadoLeitura = service.ler("update.xls")

        Assertions.assertTrue(Objects.nonNull(dadoLeitura))
        Assertions.assertEquals(1, dadoLeitura.size)
        Assertions.assertTrue(dadoLeitura[0][0].title)
        Assertions.assertEquals("teste", dadoLeitura[0][0].value)
        Assertions.assertEquals(0, dadoLeitura[0][0].position)
        Assertions.assertEquals("Isso e um teste unitario", dadoLeitura[0][1].value)

    }

    @Test
    @Order(4)
    fun deveMoverUmaPlanilha() {
        val service = ExcelService()

        service.mover("teste.xls", "movido.xls")

        val planilha = File("movido.xls")

        Assertions.assertTrue(planilha.exists())
    }

    @Test
    @Order(5)
    fun deveDeletarPlanilhas() {
        val service = ExcelService()
        val dado: DadoTeste = DadoTeste("teste", "Opa", 0)
        val lista : List<Excel> = mutableListOf(dado)
        val dadosPlanilha : List<List<Excel>> = mutableListOf(lista)

        service.novo(dadosPlanilha, "apagar.xls")

        val planilha1 = File("apagar.xls")
        Assertions.assertTrue(planilha1.exists())

        service.apagar("apagar.xls")

        val planilha = File("apagar.xls")
        Assertions.assertFalse(planilha.exists())
    }
}