package dos.gatos.poi.service;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import dos.gatos.poi.util.WorkbookConfig;
import dos.gatos.poi.util.WorkbookConfig.StyleConfig;
import dos.gatos.poi.util.WorkbookConfig.WorkbookConfigBuilder;
import dos.gatos.poi.util.WorkbookGenerator;
import lombok.Builder;
import lombok.Value;

@RestController
@RequestMapping("/workbooks")
public class WorkbookService {

	private static final WorkbookConfig<MyModel> CONFIG;
	static {
		WorkbookConfigBuilder<MyModel> config = WorkbookConfig.builder(MyModel.class);
		config.name("Staff").sheetName("MyList").defaultHeaderStyle().defaultBodyStyle();
		config.numberCol("ID", MyModel::getId);
		config.stringCol("Name", MyModel::getName);
		config.dateCol("BirthDate", MyModel::getDate, StyleConfig.builder().dataFormat("yyyy-mm-dd").build());
		CONFIG = config.build();
	}

	@GetMapping(produces = "application/vnd.ms-excel")
	public XSSFWorkbook getExcel() {
		List<MyModel> data = new ArrayList<>();
		data.add(new MyModel(1, "Dan", LocalDate.of(2019, 11, 1)));
		data.add(new MyModel(2, "Chris", LocalDate.of(2019, 11, 2)));
		data.add(new MyModel(3, "Peter", LocalDate.of(2019, 11, 3)));
		data.add(new MyModel(4, "Kate", LocalDate.of(2019, 11, 4)));
		data.add(new MyModel(5, "Jeff", LocalDate.of(2019, 11, 5)));
		return WorkbookGenerator.of(CONFIG, data);
	}

	@Value
	@Builder
	public static class MyModel {

		Integer id;
		String name;
		LocalDate date;

	}
}
