package br.net.silva.daniel.ex.plan.model

interface Excel {
    fun title() : String
    fun value() : Any
    fun positionTitle() : Int
    fun positionValue() : Int
}