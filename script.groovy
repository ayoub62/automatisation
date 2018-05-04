import groovy.sql.Sql
import java.sql.DriverManager

import java.io.File
import java.io.IOException
import jxl.*
import jxl.write.*

// ------------------------------- PARAMETRAGE -------------------------------
def serviceID = 5
def balisePrin = "customerEmailAddress"
def serviceName = "getCustomerEmailAddressList"

// ------------------------------- CONNEXION A LA BASE DE DONNEES AMPLITUDE -------------------------------
com.eviware.soapui.support.GroovyUtils.registerJdbcDriver( "oracle.jdbc.driver.OracleDriver")



sql = Sql.newInstance(dbUrl, dbUser, dbPassword, dbDriver)

// ------------------------------- RECUPERER LA REQUETE COMPLETE -------------------------------
import groovyx.net.http.RESTClient

def client = new RESTClient( 'http://localhost:8080' )
def resp = client.get( path : 'piste2/'+serviceID )

Scanner scanner = new Scanner(resp.getData()).useDelimiter("\\A")
String req = scanner.next()

// ------------------------------- RECUPERER TOUS LES CHAMPS DU FLUX DE REPONSE -------------------------------

def testStep = testRunner.testCase.testSteps["pro"]

// ------------------------------- SENDING THE SOAP REQUEST -------------------------------

testRunner.testCase.getTestStepByName("soapReq").run(testRunner,context)
def groovyUtils = new com.eviware.soapui.support.GroovyUtils( context )
def responseHolder = groovyUtils.getXmlHolder( testRunner.testCase.testSteps["soapReq"].testRequest.response.responseContent )


WritableWorkbook workbook1 = Workbook.createWorkbook(new File("d:\\Profiles\\aaitlahcen\\Desktop\\PERSO\\"+serviceName+".xls"))
WritableSheet sheet1 = workbook1.createSheet("Rapport", 0)

WritableFont cellFont = new WritableFont(WritableFont.TIMES, 12);
cellFont.setColour(Colour.WHITE);

WritableCellFormat cellFormatV = new WritableCellFormat(cellFont);
cellFormatV.setBackground(Colour.GREEN);
cellFormatV.setBorder(Border.ALL, BorderLineStyle.THIN);

WritableCellFormat cellFormatNV = new WritableCellFormat(cellFont);
cellFormatNV.setBackground(Colour.RED);
cellFormatNV.setBorder(Border.ALL, BorderLineStyle.THIN);
sheet1.addCell(new Label(0, 0, "Le statut de la réponse"));
sheet1.addCell(new Label(1, 0, responseHolder.getNodeValue("//fjs1:statusCode")));
int x = 1
int j = 1;
sql.eachRow(req){row ->
	def champsClient = client.get( path : 'xpath/'+serviceID )
	for(String s : champsClient.getData()){
		String[] split = s.split(":")[1].split("_")
		String memeTable = s.split(":")[2]
		String calculated = s.split(":")[3]
		def xpath = "//fjs1:"+balisePrin+"[$x]";
		for(int i = split.length - 1; i >= 0;i--){
			xpath += "//fjs1:"+split[i]
		}
		String name = s.split(":")[1]
		
		String p = "champ/"+serviceID+"/"+name
		def cClient = client.get( path : p)
		def resultFromRest = new Scanner(cClient.getData()).useDelimiter("\\A").next()
		def expected
		def observed = responseHolder.getNodeValue(xpath).toString().trim()
		if(memeTable.equals("true") && calculated.equals("false")){
			expected = row."$resultFromRest".toString().trim()
		}
		else{
			String[] splitReq = resultFromRest.split("\\*")
			String param = splitReq[0].split(" ")[1]
			String tagValue = testRunner.testCase.testSteps["proResponse"].getPropertyValue(splitReq[1])
			//log.info splitReq[0] + " : "+tagValue
			completeReq = splitReq[0].replaceAll("\\?",tagValue)
			log.info name+" : "+completeReq
			def myReq = sql.firstRow(completeReq)
			if(calculated.equals("false"))
				expected = myReq."$param".toString().trim()
			else{
				expected = myReq."$name".toString().trim()
			}
		}
		//log.info expected+" : "+observed
		sheet1.addCell(new Label(0, j, name));
		if(name.startsWith("date")){
			if(expected != "null")
				expected = expected.substring(0,10);
			if(observed != "null")
				observed = observed.substring(0,10);
		}
		if(name.startsWith("amount") || name.contains("Rate") || name.contains("Capital")){
			expected = Double.parseDouble(expected+"");
			observed = Double.parseDouble(observed+"");
			expected = String.format("%f",expected);
			observed = String.format("%f",observed);
		}
		if(expected == "null")
			expected = ""
		if(observed == "null")
			observed = ""
		sheet1.addCell(new Label(1, j, expected));
		sheet1.addCell(new Label(2, j, observed));
		if(expected == observed){
			sheet1.addCell(new Label(3, j, "PASS",cellFormatV));
		}
		else{
			sheet1.addCell(new Label(3, j, "FAIL",cellFormatNV));
		}

		j++;
		testRunner.testCase.testSteps["proResponse"].setPropertyValue(name,expected)
	}
	x++;
}

workbook1.write()
workbook1.close()