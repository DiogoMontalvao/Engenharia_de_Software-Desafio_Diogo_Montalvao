
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
import java.io.InputStreamReader

object SheetsQuickstart {

    private const val APPLICATION_NAME = "Desafio Diogo Montalvão"
    private val JSON_FACTORY = GsonFactory.getDefaultInstance()
    private const val TOKENS_DIRECTORY_PATH = "tokens"

    private val SCOPES = listOf(SheetsScopes.SPREADSHEETS)
    private const val CREDENTIALS_FILE_PATH = "/credentials.json"

    @JvmStatic
    fun main(args: Array<String>) {
        val HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport()
        val spreadsheetId = "1TDNQ2ktrrtx_oiXkEwIYcU_kuYY9dHMqvLoMWQfX86g"

        val range = "engenharia_de_software!A4:H"
        val service = Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
            .setApplicationName(APPLICATION_NAME)
            .build()
        val response = service.spreadsheets().values()
            .get(spreadsheetId, range)
            .execute()

        val values = response.getValues()

        if (values.isNullOrEmpty()) {
            println("Dados não encontrados")
        } else {
            for (coluna in values) {
                println(
                    """
                Matrícula: ${coluna[0]}
                Aluno: ${coluna[1]}
                Faltas: ${coluna[2]}
                
                """.trimIndent()
                )
            }
        }
    }

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
}




