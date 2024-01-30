import com.google.api.client.auth.oauth2.Credential
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport
import com.google.api.client.http.javanet.NetHttpTransport
import com.google.api.client.json.gson.GsonFactory
import com.google.api.client.util.store.FileDataStoreFactory
import com.google.api.services.sheets.v4.Sheets
import com.google.api.services.sheets.v4.SheetsScopes
import com.google.api.services.sheets.v4.model.UpdateValuesResponse
import com.google.api.services.sheets.v4.model.ValueRange
import java.io.InputStreamReader
import java.math.BigDecimal
import java.math.RoundingMode

private object SheetsQuickstart {

    private const val APPLICATION_NAME = "Desafio Diogo Montalvão"
    private val JSON_FACTORY = GsonFactory.getDefaultInstance()
    private const val TOKENS_DIRECTORY_PATH = "tokens/path"

    private val SCOPES = listOf(SheetsScopes.SPREADSHEETS)
    private const val CREDENTIALS_FILE_PATH = "/credentials.json"

    @JvmStatic
    fun main(args: Array<String>) {
        val HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport()
        val spreadsheetId = "1TDNQ2ktrrtx_oiXkEwIYcU_kuYY9dHMqvLoMWQfX86g"
        val range = "engenharia_de_software!A4:H"

        val service = Sheets
            .Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
            .setApplicationName(APPLICATION_NAME)
            .build()

        val valuesSpreadsheet = getValues(spreadsheetId, range, service).getValues()

        var currentRow = 4
        for (row in valuesSpreadsheet) {
            val aluno = row[1].toString()
            val faltas = row[2].toString().toDouble()

            val prova1 = row[3].toString().toBigDecimal()
            val prova2 = row[4].toString().toBigDecimal()
            val prova3 = row[5].toString().toBigDecimal()
            val totalProvas = (prova1 + prova2 + prova3).setScale(1)

            val situacao = updateSituacao(faltas, totalProvas)
            val notaAprovacaoFinal = updateNotaAprovacaoFinal(situacao, totalProvas)

            updateValues(
                spreadsheetId,
                "engenharia_de_software!G$currentRow:H$currentRow",
                valueInputOption = "USER_ENTERED",
                listOf(listOf(situacao, notaAprovacaoFinal)),
                service
            )

            println("Atualizando - Aluno: $aluno, situação: $situacao")

            currentRow++
        }
    }

    /**
     * Creates an authorized Credential object.
     *
     * @param HTTP_TRANSPORT - The network HTTP Transport.
     * @return An authorized Credential object.
     */
    private fun getCredentials(HTTP_TRANSPORT: NetHttpTransport): Credential {
        val `in` = SheetsQuickstart.javaClass.getResourceAsStream(CREDENTIALS_FILE_PATH)
        val clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, `in`?.let { InputStreamReader(it) })

        val flow = GoogleAuthorizationCodeFlow
            .Builder(HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
            .setDataStoreFactory(FileDataStoreFactory(java.io.File(TOKENS_DIRECTORY_PATH)))
            .setAccessType("online")
            .build()

        val receiver = LocalServerReceiver
            .Builder()
            .setPort(8888)
            .build()

        return AuthorizationCodeInstalledApp(flow, receiver).authorize("user")
    }

    /**
     * Returns a range of values from a spreadsheet.
     *
     * @param spreadsheetId - Id of the spreadsheet.
     * @param range         - Range of cells of the spreadsheet.
     * @param service       - Create the sheets API client.
     * @return Values in the range.
     */
    private fun getValues(spreadsheetId: String, range: String, service: Sheets): ValueRange {
        val result = service
            .spreadsheets()
            .values()
            .get(spreadsheetId, range)
            .execute()

        return result
    }

    /**
     * Sets values in a range of a spreadsheet.
     *
     * @param spreadsheetId    - ID of the spreadsheet.
     * @param range            - Range of cells of the spreadsheet.
     * @param valueInputOption - Determines how input data should be interpreted.
     * @param values           - List of rows of values to input.
     * @param service          - Create the sheets API client.
     * @return Spreadsheet with updated values.
     */
    private fun updateValues(
        spreadsheetId: String,
        range: String,
        valueInputOption: String,
        values: List<List<Any>>,
        service: Sheets
    ): UpdateValuesResponse {
        val body = ValueRange().setValues(values)

        val result = service
            .spreadsheets()
            .values()
            .update(spreadsheetId, range, body)
            .setValueInputOption(valueInputOption)
            .execute()

        return result
    }

    /**
     * Updates the student's "Situação".
     * @param faltas      - Student's absences.
     * @param totalProvas - Total value of the student's exams.
     * @return Student's "Situação".
     */
    private fun updateSituacao(faltas: Double, totalProvas: BigDecimal): String {
        var situacao: String
        val totalAulasSemestre = 60
        val indiceFaltas = faltas / totalAulasSemestre
        val mediaProvas = calculaMediaProvas(totalProvas)

        when {
            mediaProvas < 5 -> situacao = "Reprovado por Nota"
            mediaProvas in 5..6 -> situacao = "Exame Final"
            else -> situacao = "Aprovado"
        }

        if (indiceFaltas > 0.25) situacao = "Reprovado por Falta"

        return situacao
    }

    /**
     * Updates the student's "Nota para Aprovação Final".
     * @param situacao    - Student's "Situação".
     * @param totalProvas - Total value of the student's exams.
     * @return "Nota para Aprovação Final" value.
     */
    private fun updateNotaAprovacaoFinal(situacao: String, totalProvas: BigDecimal): Int {
        val notaAprovacaoFinal: Int
        val mediaProvas = calculaMediaProvas(totalProvas)

        notaAprovacaoFinal =
            if (situacao != "Exame Final") {
                0
            } else {
                10 - mediaProvas
            }

        return notaAprovacaoFinal
    }

    /**
     * Calculates the average value of the student's exams.
     * @param totalProvas - Total value of the student's exams.
     * @return Average value of the student's exams.
     */
    private fun calculaMediaProvas(totalProvas: BigDecimal): Int {
        return totalProvas
            .div("3.0".toBigDecimal())
            .div("10.0".toBigDecimal())
            .setScale(0, RoundingMode.HALF_UP)
            .toInt()
    }
}




