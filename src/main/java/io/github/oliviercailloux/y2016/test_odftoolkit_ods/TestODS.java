package io.github.oliviercailloux.y2016.test_odftoolkit_ods;

import java.io.InputStream;

import org.odftoolkit.simple.SpreadsheetDocument;
import org.odftoolkit.simple.style.Border;
import org.odftoolkit.simple.style.StyleTypeDefinitions.CellBordersType;
import org.odftoolkit.simple.table.Cell;
import org.odftoolkit.simple.table.Table;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Simplified from <a href=
 * "https://incubator.apache.org/odftoolkit/simple/demo/demo9.html">odftoolkit
 * doc</a>.
 *
 * @author Olivier Cailloux
 *
 */
public class TestODS {
	@SuppressWarnings("unused")
	private static final Logger LOGGER = LoggerFactory.getLogger(TestODS.class);

	public static void main(String[] args) throws Exception {
		final TestODS testODS = new TestODS();
		testODS.readBordersIAE();
	}

	public void readBordersIAE() throws Exception {
		LOGGER.debug("Reading.");
		try (InputStream inputStream = TestODS.class.getResourceAsStream("IAE.ods");
				SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.loadDocument(inputStream)) {
			LOGGER.debug("Read.");
			final Table table = spreadsheetDoc.getTableByName("Table");
			final Cell a1Cell = table.getCellByPosition("A1");
			LOGGER.info("Found: {}.", a1Cell.getDisplayText());
			final Cell borderCell = table.getCellByPosition("F4");
			final Border left = borderCell.getStyleHandler().getBorder(CellBordersType.LEFT);
			LOGGER.info("Border left: {}.", left);
			final Border borderBottomLeft = borderCell.getStyleHandler().getBorder(CellBordersType.DIAGONALBLTR);
			LOGGER.info("Border bottom left: {}.", borderBottomLeft);
			final Border borderTopLeft = borderCell.getStyleHandler().getBorder(CellBordersType.DIAGONALTLBR);
			LOGGER.info("Border top left: {}.", borderTopLeft);
		}
	}

	public void readBordersNFE() throws Exception {
		LOGGER.debug("Reading.");
		try (InputStream inputStream = TestODS.class.getResourceAsStream("NFE.ods");
				SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.loadDocument(inputStream)) {
			LOGGER.debug("Read.");
			final Table table = spreadsheetDoc.getTableByName("L3_Informatique");
			final Cell a1Cell = table.getCellByPosition("A1");
			LOGGER.info("Found: {}.", a1Cell.getDisplayText());
			final Cell borderCell = table.getCellByPosition("F4");
			final Border left = borderCell.getStyleHandler().getBorder(CellBordersType.LEFT);
			LOGGER.info("Border left: {}.", left);
			final Border borderBottomLeft = borderCell.getStyleHandler().getBorder(CellBordersType.DIAGONALBLTR);
			LOGGER.info("Border bottom left: {}.", borderBottomLeft);
			final Border borderTopLeft = borderCell.getStyleHandler().getBorder(CellBordersType.DIAGONALTLBR);
			LOGGER.info("Border top left: {}.", borderTopLeft);
		}
	}

	public void readBordersNPE() throws Exception {
		LOGGER.debug("Reading.");
		try (InputStream inputStream = TestODS.class.getResourceAsStream("NPE.ods");
				SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.loadDocument(inputStream)) {
			LOGGER.debug("Read.");
			final Table table = spreadsheetDoc.getTableByName("Table");
			final Cell a1Cell = table.getCellByPosition("A1");
			LOGGER.info("Found: {}.", a1Cell.getDisplayText());
			final Cell borderCell = table.getCellByPosition("F4");
			final Border left = borderCell.getStyleHandler().getBorder(CellBordersType.LEFT);
			LOGGER.info("Border left: {}.", left);
			final Border borderBottomLeft = borderCell.getStyleHandler().getBorder(CellBordersType.DIAGONALBLTR);
			LOGGER.info("Border bottom left: {}.", borderBottomLeft);
			final Border borderTopLeft = borderCell.getStyleHandler().getBorder(CellBordersType.DIAGONALTLBR);
			LOGGER.info("Border top left: {}.", borderTopLeft);
		}
	}
}
