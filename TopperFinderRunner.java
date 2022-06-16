package JSON;

import java.io.IOException;

import org.json.simple.parser.ParseException;

public class TopperFinderRunner {
	public static void main(String[] args) throws IOException, ParseException {
		TopperFinder oFinder = new TopperFinder();
		oFinder.writeToJsonFile();
		oFinder.findTheTopperStudent();
		oFinder.writetoExcel();
	}
}
