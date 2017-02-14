# Hkwg Converter 
Es handelt sich hier um eine .NET Kommandozeilenanwendung, zum Konvertieren von CSV Datei in ein spezielles
Excel-Format und zurück.

## Einstellungen

Um das Programm laufen zu lassen, müssen einige Pfade für die Zeillandschaft eingestellt werden.
Diese Einstellungen befinden sich in der Datei HkwgConverter.exe.config.

### LogFile-Verzeichnis

Das Programm protokolliert seine Aktivität in einem Logfile. Hierfür sollte ein Ordner angelegt werden.

	
	 <log4net>    
	    <appender name="RollingFileAppender" type="log4net.Appender.RollingFileAppender">
	      <file type="log4net.Util.PatternString" value="C:\temp\logs\hkwg.log" />
	      <appendToFile value="true" />
		...
	 </log4net> 
	

### Inbound Verzeichnisse

Ankommende CSV Dateien werden im "InboundWatchFolder" erwartet und dann in die entsprechenden Unterordner verschoben.
	
	  <applicationSettings>
	        <HkwgConverter.Settings>
	            <setting name="InboundWatchFolder" serializeAs="String">
	                <value>D:\FileStore\Inbound\</value>
	            </setting>
	            <setting name="InboundSuccessFolder" serializeAs="String">
	                <value>D:\FileStore\Inbound\OK\</value>
	            </setting>
	            <setting name="InboundErrorFolder" serializeAs="String">
	                <value>D:\FileStore\Inbound\NOK\</value>
	            </setting>                      
	        </HkwgConverter.Settings>
	    </applicationSettings>


Die erzeugten XSLX Dateien werden im "InboundWatchFolder" abgelegt und müssen durch einen Folgeprozess weiterverarbeitet werden.

### Outbound Verzeichnisse

Ankommende XSLX Dateien werden im "OutboudWatchFolder" erwartet und dann in die entsprechenden Unterordner verschoben.

	  <applicationSettings>
	        <HkwgConverter.Settings>
	             <setting name="OutboundWatchFolder" serializeAs="String">
	                <value>D:\FileStore\Outbound\</value>
	            </setting>
	            <setting name="OutboundSuccessFolder" serializeAs="String">
	                <value>D:\FileStore\Outbound\OK\</value>
	            </setting>
	            <setting name="OutboundErrorFolder" serializeAs="String">
	                <value>D:\FileStore\Outbound\NOK\</value>
	            </setting>
	        </HkwgConverter.Settings>
	    </applicationSettings>


### AppData Verzeichnis

Hier protokolliert die Anwendung einige Informationen über die verarbeiteten Dateien.

	  <applicationSettings>
	        <HkwgConverter.Settings>
				<setting name="AppDataFolder" serializeAs="String">
	                <value>D:\FileStore\AppData\</value>
	            </setting>
	        </HkwgConverter.Settings>
	    </applicationSettings>


