package JSON;

import java.io.IOException;

import org.json.simple.parser.ParseException;

public class JsonWriter {

	public static void main(String[] args) throws ParseException, IOException {
		JsonFile oFile = new JsonFile();
		oFile.jsonwriter();
		oFile.jsonReader();

	}

}
