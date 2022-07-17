package com.leitorplanilha

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File
import java.io.FileInputStream

fun main() {
    val corretoras = readSpreadSheet("C://Pasta1.xlsx")
    val meuCsv = File("C://meucsv.csv").readLines()
    val conteudoCsv = meuCsv.map { it.split(";") }

    for (crt in corretoras) {
        val linhaCsv = conteudoCsv.firstOrNull { it[2].canBeDouble() && crt.codigo == it[2].toDouble() }
        crt.cnpj = linhaCsv?.get(2) ?: "Não encontrado"
    }

    corretoras.forEach {
        println("codigo ${it.codigo}")
        println("nome ${it.nome}")
        println("cnpj ${it.cnpj}")
        println()
    }

    writeCsv(corretoras)
    writeSql(corretoras)
}

data class Corretora(
    val codigo: Double,
    val nome: String,
    var cnpj: String? = null,
){
    override fun toString():String =
        "$codigo;$nome;$cnpj"
}

fun readSpreadSheet(filePath: String): MutableList<Corretora> {
    var arquivo: FileInputStream? = null
    var pastaDeTrabalho: Workbook? = null
    try {
        arquivo = FileInputStream(filePath)
        pastaDeTrabalho = WorkbookFactory.create(arquivo)
        val planilha = pastaDeTrabalho.getSheet("Planilha1")

        val linhas = planilha.toMutableList()
        val corretoras = mutableListOf<Corretora>()
        linhas.removeFirst()

        linhas.forEach {
            if (it.getCell(0).cellType == CellType.NUMERIC) {
                val corretora = Corretora(
                    it.getCell(0).numericCellValue,
                    it.getCell(2).toString()
                )
                corretoras.add(corretora)
            }
        }
        return corretoras
    } catch (ex: Exception) {
        ex.printStackTrace()
    } finally {
        pastaDeTrabalho?.close()
        arquivo?.close()
    }
    return mutableListOf()
}

fun String.canBeDouble(): Boolean {
    return try {
        this.toDouble()
        true
    } catch (ex: NumberFormatException) {
        false
    }
}

fun writeCsv(corretoras: List<Corretora>) {
    val strBuilder=StringBuilder()
    corretoras.forEach{
        println(it.toString())
        strBuilder.appendLine(it.toString())
    }
    File("minhasCorretoras.csv").writeText(strBuilder.toString())
}

fun writeSql(corretoras: List<Corretora>){
    val strBuilder=StringBuilder()
    corretoras.forEach{
        val sql="INSERT INTO table() VALUES (${it.codigo},${it.nome})"
        strBuilder.appendLine(it.toString())
    }
    File("minhasCorretoras.sql").writeText(strBuilder.toString())
}