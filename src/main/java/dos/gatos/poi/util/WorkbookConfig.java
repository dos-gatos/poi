package dos.gatos.poi.util;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import lombok.Builder;
import lombok.EqualsAndHashCode;
import lombok.NonNull;
import lombok.ToString;
import lombok.Value;

@Value
@ToString
public class WorkbookConfig<T> {

	Class<T> source;
	String name;
	String sheetName;
	StyleConfig headerStyle;
	StyleConfig bodyStyle;
	List<ColumnConfig<T, ?>> columns;

	public Set<StyleConfig> getStyleSet() {
		Set<StyleConfig> styleSet = new HashSet<>();
		styleSet.add(headerStyle);
		styleSet.add(bodyStyle);
		styleSet.addAll(columns.stream().map(ColumnConfig::getStyle).collect(Collectors.toSet()));
		return styleSet;
	}

	public Set<FontConfig> getFontSet() {
		Set<FontConfig> fontSet = new HashSet<>();
		headerStyle.getFontConfig().ifPresent(fontSet::add);
		bodyStyle.getFontConfig().ifPresent(fontSet::add);
		columns.stream().forEach(column -> column.getStyle().getFontConfig().ifPresent(fontSet::add));
		return fontSet;
	}

	/*** BUILDER ***/

	public static <T> WorkbookConfigBuilder<T> builder(Class<T> source) {
		return new WorkbookConfigBuilder<>(source);
	}

	public interface NameStep<T> {
		SheetNameStep<T> name(String name);
	}

	public interface SheetNameStep<T> {
		HeaderStyleStep<T> sheetName(String sheetName);
	}

	public interface HeaderStyleStep<T> {
		BodyStyleStep<T> headerStyle(StyleConfig headerStyle);
		BodyStyleStep<T> defaultHeaderStyle();
	}

	public interface BodyStyleStep<T> {
		ColumnStep<T> bodyStyle(StyleConfig columnStyle);
		ColumnStep<T> defaultBodyStyle();
	}

	public interface ColumnStep<T> {
		ColumnStep<T> stringCol(String name, Function<T, String> mapper);
		ColumnStep<T> stringCol(String name, Function<T, String> mapper, StyleConfig style);
		ColumnStep<T> numberCol(String name, Function<T, Number> mapper);
		ColumnStep<T> numberCol(String name, Function<T, Number> mapper, StyleConfig style);
		ColumnStep<T> dateCol(String name, Function<T, LocalDate> mapper);
		ColumnStep<T> dateCol(String name, Function<T, LocalDate> mapper, StyleConfig style);
		ColumnStep<T> datetimeCol(String name, Function<T, LocalDateTime> mapper);
		ColumnStep<T> datetimeCol(String name, Function<T, LocalDateTime> mapper, StyleConfig style);
		ColumnStep<T> booleanCol(String name, Function<T, Boolean> mapper);
		ColumnStep<T> booleanCol(String name, Function<T, Boolean> mapper, StyleConfig style);
		WorkbookConfig<T> build();
	}

	public static class WorkbookConfigBuilder<T> implements NameStep<T>, SheetNameStep<T>, HeaderStyleStep<T>, BodyStyleStep<T>, ColumnStep<T> {

		private Class<T> source;
		private String name;
		private String sheetName;
		private StyleConfig headerStyle;
		private StyleConfig bodyStyle;
		private List<ColumnConfig<T, ?>> columns = new ArrayList<>();

		private WorkbookConfigBuilder(Class<T> source) {
			this.source = source;
		}

		@Override
		public SheetNameStep<T> name(String name) {
			this.name = Optional.ofNullable(name).orElse("Workbook1");
			return this;
		}

		@Override
		public HeaderStyleStep<T> sheetName(String sheetName) {
			this.sheetName = Optional.ofNullable(sheetName).orElse("Sheet1");
			return this;
		}

		@Override
		public BodyStyleStep<T> headerStyle(StyleConfig style) {
			this.headerStyle = Optional.ofNullable(style).orElse(StyleConfig.DEFAULT_HEADER);
			return this;
		}

		@Override
		public BodyStyleStep<T> defaultHeaderStyle() {
			return headerStyle(null);
		}

		@Override
		public ColumnStep<T> bodyStyle(StyleConfig style) {
			this.bodyStyle = Optional.ofNullable(style).orElse(StyleConfig.DEFAULT_BODY);
			return this;
		}

		@Override
		public ColumnStep<T> defaultBodyStyle() {
			return bodyStyle(null);
		}

		@Override
		public ColumnStep<T> stringCol(String name, Function<T, String> mapper) {
			return col(name, String.class, mapper);
		}

		@Override
		public ColumnStep<T> stringCol(String name, Function<T, String> mapper, StyleConfig style) {
			return col(name, String.class, mapper, style);
		}

		@Override
		public ColumnStep<T> numberCol(String name, Function<T, Number> mapper) {
			return col(name, Double.class, mapper, Number::doubleValue);
		}

		@Override
		public ColumnStep<T> numberCol(String name, Function<T, Number> mapper, StyleConfig style) {
			return col(name, Double.class, mapper, Number::doubleValue, style);
		}

		@Override
		public ColumnStep<T> dateCol(String name, Function<T, LocalDate> mapper) {
			final Function<LocalDate, Date> mapper2 = ld -> Date.from(ld.atStartOfDay().atZone(ZoneId.systemDefault()).toInstant());
			return col(name, Date.class, mapper, mapper2);
		}

		@Override
		public ColumnStep<T> dateCol(String name, Function<T, LocalDate> mapper, StyleConfig style) {
			final Function<LocalDate, Date> mapper2 = ld -> Date.from(ld.atStartOfDay().atZone(ZoneId.systemDefault()).toInstant());
			return col(name, Date.class, mapper, mapper2, style);
		}

		@Override
		public ColumnStep<T> datetimeCol(String name, Function<T, LocalDateTime> mapper) {
			final Function<LocalDateTime, Date> mapper2 = ldt -> Date.from(ldt.atZone(ZoneId.systemDefault()).toInstant());
			return col(name, Date.class, mapper, mapper2);
		}

		@Override
		public ColumnStep<T> datetimeCol(String name, Function<T, LocalDateTime> mapper, StyleConfig style) {
			final Function<LocalDateTime, Date> mapper2 = ldt -> Date.from(ldt.atZone(ZoneId.systemDefault()).toInstant());
			return col(name, Date.class, mapper, mapper2, style);
		}

		@Override
		public ColumnStep<T> booleanCol(String name, Function<T, Boolean> mapper) {
			return col(name, Boolean.class, mapper);
		}

		@Override
		public ColumnStep<T> booleanCol(String name, Function<T, Boolean> mapper, StyleConfig style) {
			return col(name, Boolean.class, mapper, style);
		}

		@Override
		public WorkbookConfig<T> build() {
			return new WorkbookConfig<>(source, name, sheetName, headerStyle, bodyStyle, columns);
		}

		private <U> ColumnStep<T> col(String name, Class<U> target, Function<T, U> mapper) {
			Objects.requireNonNull(name, "column name is required");
			Objects.requireNonNull(mapper, "column mapping is required");
			this.columns.add(ColumnConfig.builder(source, target).name(name).mapper(mapper).style(bodyStyle).build());
			return this;
		}

		private <U, V> ColumnStep<T> col(String name, Class<V> target, Function<T, U> mapper1, Function<U, V> mapper2) {
			Objects.requireNonNull(name, "column name is required");
			Objects.requireNonNull(mapper1, "column mapping is required");
			Objects.requireNonNull(mapper2, "column mapping is required");
			this.columns.add(ColumnConfig.builder(source, target).name(name).mapper(mapper2.compose(mapper1)).style(bodyStyle).build());
			return this;
		}

		private <U> ColumnStep<T> col(String name, Class<U> target, Function<T, U> mapper, StyleConfig style) {
			Objects.requireNonNull(name, "column name is required");
			Objects.requireNonNull(mapper, "column mapping is required");
			style = Optional.ofNullable(style).map(s -> StyleConfig.builder(this.bodyStyle, s).build()).orElse(this.bodyStyle);
			this.columns.add(ColumnConfig.builder(source, target).name(name).mapper(mapper).style(style).build());
			return this;
		}

		private <U, V> ColumnStep<T> col(String name, Class<V> target, Function<T, U> mapper1, Function<U, V> mapper2, StyleConfig style) {
			Objects.requireNonNull(name, "column name is required");
			Objects.requireNonNull(mapper1, "column mapping is required");
			Objects.requireNonNull(mapper2, "column mapping is required");
			style = Optional.ofNullable(style).map(s -> StyleConfig.builder(bodyStyle, s).build()).orElse(bodyStyle);
			this.columns.add(ColumnConfig.builder(source, target).name(name).mapper(mapper2.compose(mapper1)).style(style).build());
			return this;
		}
	}

	@Value
	@Builder
	@EqualsAndHashCode
	public static class ColumnConfig<T, U> {

		@NonNull
		String name;
		@NonNull
		StyleConfig style;
		@NonNull
		Class<T> source;
		@NonNull
		Class<U> target;
		@NonNull
		Function<T, U> mapper;

		public static <T, U> ColumnConfigBuilder<T, U> builder() {
			return new ColumnConfigBuilder<>();
		}

		public static <T, U> ColumnConfigBuilder<T, U> builder(Class<T> source, Class<U> target) {
			ColumnConfigBuilder<T, U> builder = builder();
			builder.source(source);
			builder.target(target);
			return builder;
		}
	}

	@Value
	@Builder
	@EqualsAndHashCode
	public static class StyleConfig {

		public static final StyleConfig DEFAULT_HEADER;
		static {
			StyleConfigBuilder builder = StyleConfig.builder();
			builder.fillPattern(FillPatternType.SOLID_FOREGROUND);
			builder.fillColor(0x000000);
			builder.horizontalAlignment(HorizontalAlignment.CENTER);
			builder.fontConfig(FontConfig.DEFAULT_HEADER);
			DEFAULT_HEADER = builder.build();
		}
		public static final StyleConfig DEFAULT_BODY = StyleConfig.builder().build();

		FillPatternType fillPattern;
		Integer fillColor;
		HorizontalAlignment horizontalAlignment;
		VerticalAlignment verticalAlignment;
		FontConfig fontConfig;
		String dataFormat;

		public Optional<FillPatternType> getFillPattern() {
			return Optional.ofNullable(fillPattern);
		}

		public Optional<Integer> getFillColor() {
			return Optional.ofNullable(fillColor);
		}

		public Optional<HorizontalAlignment> getHorizontalAlignment() {
			return Optional.ofNullable(horizontalAlignment);
		}

		public Optional<VerticalAlignment> getVerticalAlignment() {
			return Optional.ofNullable(verticalAlignment);
		}

		public Optional<FontConfig> getFontConfig() {
			return Optional.ofNullable(fontConfig);
		}

		public Optional<String> getDataFormat() {
			return Optional.ofNullable(dataFormat);
		}

		public static StyleConfigBuilder builder() {
			return new StyleConfigBuilder();
		}

		public static StyleConfigBuilder builder(StyleConfig... styles) {
			StyleConfigBuilder builder = builder();
			Arrays.stream(styles).forEach(style -> {
				style.getFillPattern().ifPresent(builder::fillPattern);
				style.getFillColor().ifPresent(builder::fillColor);
				style.getHorizontalAlignment().ifPresent(builder::horizontalAlignment);
				style.getVerticalAlignment().ifPresent(builder::verticalAlignment);
				style.getFontConfig().ifPresent(builder::fontConfig);
				style.getDataFormat().ifPresent(builder::dataFormat);
			});
			return builder;
		}
	}

	@Value
	@Builder
	@EqualsAndHashCode
	public static class FontConfig {

		public static final FontConfig DEFAULT_HEADER;
		static {
			FontConfigBuilder builder = FontConfig.builder();
			builder.color(0xFFFFFF);
			DEFAULT_HEADER = builder.build();
		}
		public static final FontConfig DEFAULT_BODY = FontConfig.builder().build();

		String name;
		Short size;
		Integer color;

		public Optional<String> getName() {
			return Optional.ofNullable(name);
		}

		public Optional<Short> getSize() {
			return Optional.ofNullable(size);
		}

		public Optional<Integer> getColor() {
			return Optional.ofNullable(color);
		}

		public static FontConfigBuilder builder() {
			return new FontConfigBuilder();
		}

		public static FontConfigBuilder builder(FontConfig fontConfig) {
			FontConfigBuilder builder = builder();
			fontConfig.getName().ifPresent(builder::name);
			fontConfig.getSize().ifPresent(builder::size);
			fontConfig.getColor().ifPresent(builder::color);
			return builder;
		}
	}

}