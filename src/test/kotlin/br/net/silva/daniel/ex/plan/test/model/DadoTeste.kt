package br.net.silva.daniel.ex.plan.test.model

import br.net.silva.daniel.ex.plan.model.Excel

class DadoTeste(var titulo: String, var valor : String, var posicao : Int) : Excel  {

    override fun title(): String {
        return titulo
    }

    override fun value(): Any {
        return valor
    }

    override fun position(): Int {
        return posicao
    }
}