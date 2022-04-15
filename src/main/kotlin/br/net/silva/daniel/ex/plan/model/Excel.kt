package br.net.silva.daniel.ex.plan.model

interface Excel {
    fun title() : Boolean
    fun value() : Any
    fun position() : Int
    fun addTitle(title : String)
    fun addValue(value : Any)
    fun addPosition(position : Int)
}