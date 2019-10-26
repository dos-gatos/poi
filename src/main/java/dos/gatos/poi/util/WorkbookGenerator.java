package dos.gatos.poi.util;

import java.awt.Color;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import dos.gatos.poi.util.WorkbookConfig.ColumnConfig;
import dos.gatos.poi.util.WorkbookConfig.FontConfig;
import dos.gatos.poi.util.WorkbookConfig.StyleConfig;

public class WorkbookGenerator<T> {


	private final WorkbookConfig<T> config;
	private final List<T> data;
	private XSSFWorkbook workbook;
	private Map<FontConfig, XSSFFont> fonts;
	private Map<StyleConfig, XSSFCellStyle> styles;
	private XSSFSheet sheet;


	public static <T> XSSFWorkbook of(WorkbookConfig<T> wbConfig, List<T> data) {
		return new WorkbookGenerator<T>(wbConfig, data).generate();
	}

	private WorkbookGenerator(WorkbookConfig<T> config, List<T> data) {
		this.config = config;
		this.data = data;
	}

	private XSSFWorkbook generate() {
		generateWorkbook();
		writeHeader();
		writeData();
		cleanup();
		return workbook;
	}
	
	private void generateWorkbook() {
		workbook = new XSSFWorkbook();
		workbook.getProperties().getCoreProperties().setTitle(config.getName().concat(".xlsx"));
		fonts = config.getFontSet().stream().collect(Collectors.toMap(Function.identity(), fc -> {
			XSSFFont font = workbook.createFont();
			fc.getName().ifPresent(font::setFontName);
			fc.getSize().ifPresent(font::setFontHeightInPoints);
			fc.getColor().ifPresent(c -> font.setColor(new XSSFColor(new Color(c), null)));
			return font;
		}));
		styles = config.getStyleSet().stream().collect(Collectors.toMap(Function.identity(), sc -> {
			XSSFCellStyle style = workbook.createCellStyle();
			sc.getFillPattern().ifPresent(style::setFillPattern);
			sc.getFillColor().ifPresent(c -> style.setFillForegroundColor(new XSSFColor(new Color(c), null)));
			sc.getHorizontalAlignment().ifPresent(style::setAlignment);
			sc.getVerticalAlignment().ifPresent(style::setVerticalAlignment);
			sc.getFontConfig().ifPresent(fc -> style.setFont(fonts.get(fc)));
			// TODO create format registry
			sc.getDataFormat().ifPresent(df -> style.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(df)));
			return style;
		}));
		sheet = workbook.createSheet(config.getSheetName());
	}

	private void writeHeader() {
		XSSFRow row = sheet.createRow(0);
		IntStream.range(0, config.getColumns().size()).forEach(i -> {
			XSSFCell cell = row.createCell(i);
			cell.setCellStyle(styles.get(config.getHeaderStyle()));
			cell.setCellValue(config.getColumns().get(i).getName());
		});
		sheet.createFreezePane(0, 1);
	}

	private void writeData() {
		IntStream.range(0, data.size()).forEach(index -> writeDataRow(sheet.createRow(index + 1), data.get(index)));
	}

	private void writeDataRow(XSSFRow row, T data) {
		IntStream.range(0, config.getColumns().size()).forEach(c -> writeDataCell(row.createCell(c), config.getColumns().get(c), data));
	}

	private <U> void writeDataCell(XSSFCell cell, ColumnConfig<T, U> column, T data) {
		cell.setCellStyle(styles.get(column.getStyle()));
		Optional<U> value = Optional.ofNullable(column.getMapper().apply(data));
		if (value.isPresent()) {
			if (column.getTarget() == String.class) {
				cell.setCellValue(String.class.cast(value.get()));
			} else if (column.getTarget() == Double.class) {
				cell.setCellValue(Double.class.cast(value.get()));
			} else if (column.getTarget() == Date.class) {
				cell.setCellValue(Date.class.cast(value.get()));
			} else if (column.getTarget() == Boolean.class) {
				cell.setCellValue(Boolean.class.cast(value.get()));
			}
		}
	}
	
	private void cleanup() {
		IntStream.range(0, config.getColumns().size()).forEach(i -> {
			sheet.autoSizeColumn(i);
			if (sheet.getColumnWidthInPixels(i) < 64) {
				sheet.setColumnWidth(i, Math.round(8.43f*256) + 200);
			}
//			if (sheet.getColumnWidthInPixels(i) > 400) {
//				sheet.setColumnWidth(i, (25 * 256) + 200);
//			}
		});
	}

}
