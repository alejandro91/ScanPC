@ECHO OFF

ECHO COMENZANDO A ESCANEAR

java -jar scanPC-0.0.1.jar --spring.config.location=file:./config/application.properties

PAUSE