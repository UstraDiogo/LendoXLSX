import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class main {

	public static void main(String[] args) {

		FileInputStream fisPlanilha = null;
		
		File file = new File("C:\\Users\\diogo.ustra\\Downloads\\FILAS_NOVOS_SERVERS.xlsx");
		
		try {
			
			fisPlanilha = new FileInputStream(file);

			// cria um workbook = planilha completa com todas as abas
			XSSFWorkbook workbook = new XSSFWorkbook(fisPlanilha);

			// recupera apenas a primeira aba da planilha
			XSSFSheet sheet = workbook.getSheetAt(0);

			// retorna todas as linhas da planilha (0) aba 1
			Iterator<Row> rowIterator = sheet.iterator();

			// varre todas as linhas da planilha
			while (rowIterator.hasNext()) {

				// recebe cada linha da planilha
				Row row = rowIterator.next();

				// pegamos todas as celulas da linha
				Iterator<Cell> cellIterator = row.iterator();

				// varre todas as celulas da planilha
				while (cellIterator.hasNext()) {

					// recebe cada celula da planilha
					Cell cell = cellIterator.next();

					switch (cell.getCellType()) {

					case Cell.CELL_TYPE_STRING:
						System.out.println("Tipo string: " + cell.getStringCellValue());
						break;
					case Cell.CELL_TYPE_NUMERIC:
						System.out.println("Tipo numérico: " + cell.getNumericCellValue());
						break;

					}

				}
			}

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
