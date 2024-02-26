import de.elo.ix.client.*;

IXConnFactory connFact = new IXConnFactory( "http: I/server: 8080/ix-Archivellix" , "IX-Tutorial", "1.0");

IXConnection conn = connFact.create("fritz", "geheim" ,
computerName, null);