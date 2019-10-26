package dos.gatos.poi.util;

import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpInputMessage;
import org.springframework.http.HttpOutputMessage;
import org.springframework.http.MediaType;
import org.springframework.http.converter.HttpMessageConverter;
import org.springframework.http.converter.HttpMessageNotReadableException;
import org.springframework.http.converter.HttpMessageNotWritableException;

public class WorkbookHttpMessageConverter implements HttpMessageConverter<XSSFWorkbook> {

	private static final MediaType MEDIA_TYPE = MediaType.parseMediaType("application/vnd.ms-excel");

	@Override
	public boolean canRead(Class<?> clazz, MediaType mediaType) {
		return false;
	}

	@Override
	public boolean canWrite(Class<?> clazz, MediaType mediaType) {
		return XSSFWorkbook.class.equals(clazz) && MEDIA_TYPE.equals(mediaType);
	}

	@Override
	public List<MediaType> getSupportedMediaTypes() {
		return List.of(MEDIA_TYPE);
	}

	@Override
	public XSSFWorkbook read(Class<? extends XSSFWorkbook> clazz, HttpInputMessage inputMessage)
			throws IOException, HttpMessageNotReadableException {
		throw new UnsupportedOperationException();
	}

	@Override
	public void write(XSSFWorkbook t, MediaType contentType, HttpOutputMessage outputMessage)
			throws IOException, HttpMessageNotWritableException {
		outputMessage.getHeaders().setContentType(MEDIA_TYPE);
		outputMessage.getHeaders().setContentDisposition(ContentDisposition.builder("attachment")
				.filename(t.getProperties().getCoreProperties().getTitle()).build());
		t.write(outputMessage.getBody());
		t.close();
	}

}
