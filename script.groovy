import groovy.sql.Sql
import java.sql.DriverManager

import java.io.File
import java.io.IOException
import jxl.*
import java.util.*
import java.util.HashMap;
import java.util.Map;
import jxl.write.*

client = new RESTClient('http://localhost:8080')

// ------------------------------- PARAMETRAGE -------------------------------
campagneTest = 1

// données temporelles
map = [:]

def caminfo = client.get( path : 'campagneInfo/'+campagneTest )
def camList = caminfo.getData()
serviceID = camList.get(0) // id
def balisePrin = camList.get(1) // MainTag
serviceName = camList.get(2) // name
limit = 30;
rowID = 0;

// ------------------------------- CONNEXION A LA BASE DE DONNEES AMPLITUDE -------------------------------

com.eviware.soapui.support.GroovyUtils.registerJdbcDriver( "oracle.jdbc.driver.OracleDriver")


def dbSchema = camList.get(3)
def dbServer = camList.get(4)
def dbUser = camList.get(5)
def dbPassword = camList.get(6)
def dbDriver = 'oracle.jdbc.driver.OracleDriver'
def dbUrl = 'jdbc:oracle:thin:@' + dbServer + ':1521:' + dbSchema

sql = Sql.newInstance(dbUrl, dbUser, dbPassword, dbDriver)

// ------------------------------- RECUPERER LES CAS DE TEST -------------------------------
def cts = new ArrayList<>();
for(int i = 7;i < camList.size();i++){
	cts.add(camList.get(i))
}

import groovyx.net.http.RESTClient
for(int s = 0;s < cts.size();s++){
	def cas = cts.get(s)
	
	// ------------------------------- RECUPERER LA REQUETE COMPLETE -------------------------------
	
	def resp = client.get( path : 'piste2/'+cas )
	
	Scanner scanner = new Scanner(resp.getData()).useDelimiter("\\A")
	String req = scanner.next()
	// ------------------------------- RECUPERER TOUS LES CHAMPS DU FLUX DE REPONSE -------------------------------
	def testStep = testRunner.testCase.testSteps["pro"]
	// ------------------------------- GETTING THE XML REQUEST -------------------------------
	
	def xmlReqCl = client.get( path : 'flow/'+cas )
	
	scanner = new Scanner(xmlReqCl.getData()).useDelimiter("\\A")
	String requestXML = scanner.next()
	
	context.testCase.getTestStepAt(0).testRequest.requestContent = requestXML
	
	
	// ------------------------------- SENDING THE SOAP REQUEST -------------------------------
		
	testRunner.testCase.getTestStepByName("soapReq").run(testRunner,context)
	def groovyUtils = new com.eviware.soapui.support.GroovyUtils( context )
	responseHolder = groovyUtils.getXmlHolder( testRunner.testCase.testSteps["soapReq"].testRequest.response.responseContent )
	
	workbook1 = Workbook.createWorkbook(new File("d:\\Profiles\\aaitlahcen\\Desktop\\PERSO\\"+serviceName+"CT"+(s+1)+".xls"))
	sheet1 = workbook1.createSheet("Rapport", 0)
	
	cellFont = new WritableFont(WritableFont.TIMES, 12);
	cellFont.setColour(Colour.WHITE);
	
	cellFormatV = new WritableCellFormat(cellFont);
	cellFormatV.setBackground(Colour.GREEN);
	cellFormatV.setBorder(Border.ALL, BorderLineStyle.THIN);
	
	cellFormatNV = new WritableCellFormat(cellFont);
	cellFormatNV.setBackground(Colour.RED);
	cellFormatNV.setBorder(Border.ALL, BorderLineStyle.THIN);
	sheet1.addCell(new Label(0, 0, "Le statut de la réponse"));
	sheet1.addCell(new Label(1, 0, responseHolder.getNodeValue("//fjs1:statusCode")));
	int x = 1
	j = 1
	solve(req,1,balisePrin)
	workbook1.write()
	workbook1.close()
	
	
}


def solve(req,x,balisePrin){
		log.info "Patientez SVP ..."
		List rows = sql.rows(req)
		log.info rows.size()
		for(int k = 0;k < rows.size();k++){
			if(k == limit)
				break;
			row = rows[k]
			def champsClient = client.get( path : 'xpath/'+serviceID+'/'+balisePrin)
			for(String s : champsClient.getData()){
				
				if(s.split(":").length <= 1){
					req = s.split("\\|")[0]
					String tagValue = map.get(s.split("\\|")[1])
					log.info tagValue+" *****"
					req = req.replaceAll('\\?',tagValue)
					solve(req,1,s.split("\\|")[2])
				}
				else{
					name = s.split(":")[1]
					id = s.split(":")[0]
					
					split = s.split(":")[1].split("_")
					memeTable = s.split(":")[2]
					calculated = s.split(":")[3]
					multiple = "false"
					def xpath = "//fjs1:"+balisePrin+"[$x]";
					for(int i = split.length - 1; i >= 0;i--){
						xpath += "//fjs1:"+split[i]
					}
					log.info xpath
					String p = "champ/"+id
					def cClient = client.get( path : p)
					def resultFromRest = new Scanner(cClient.getData()).useDelimiter("\\A").next()
					expected = ""
					observed = responseHolder.getNodeValue(xpath).toString().trim()
					/////////////////
					observed = responseHolder.getNodeValue(xpath).toString().trim()
	
					if(memeTable.equals("true") && calculated.equals("false")){
						expected = row."$resultFromRest".toString().trim()
					}
	
					else{
	
						String[] splitReq = resultFromRest.split("\\|")
						String param = splitReq[0].split(" ")[1]
						String tagValue = map.get(splitReq[1])
						completeReq = splitReq[0].replaceAll("\\?",tagValue)
						def myReq = sql.firstRow(completeReq)
						if(calculated.equals("false"))
							expected = myReq."$param".toString().trim()
						else{
							expected = myReq."$name".toString().trim()
						}
	
					}
					sheet1.addCell(new Label(0, j, name));
					if(name.contains("date") || name.contains("Date")){
						if(expected != "null")
							expected = expected.substring(0,10);
						if(observed != "null")
							observed = observed.substring(0,10);
					}
					if(name.startsWith("amount") || name.contains("Rate") || name.contains("Capital")){
						if(expected != "null"){
							expected = Double.parseDouble(expected+"");
							expected = String.format("%f",expected);
						}
						if(observed != "null"){
							observed = Double.parseDouble(observed+"");
							observed = String.format("%f",observed);
						}
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
					map.put(name,expected)
				}
				log.info name
			}
			x++;
		}
	}



