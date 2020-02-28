package com.digiturtle.compiler;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageMargins;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

public class DocxBuilder {

	private XWPFDocument document;

	private XWPFParagraph p;

	private XWPFRun r;

	private String filename;

	private FileOutputStream stream;

	public DocxBuilder(String filename) {
		document = new XWPFDocument();
		p = document.createParagraph();
		r = p.createRun();
		this.filename = filename;
	}

	public void addImageToWordDocument(File imageFile) throws IOException, InvalidFormatException {
		BufferedImage bimg1 = ImageIO.read(imageFile);
		int width = bimg1.getWidth();
		int height = bimg1.getHeight();  

		String imgFile = imageFile.getName(); 
		int imgFormat = getImageFormat(imgFile);
		double imgRatio = (double) height / (double) width;
		int imageDocWidth = 468; // 7.5 inches * 72 points = 540, 6.5 inches * 72 points = 468
		int imageDocHeight = (int) ((double) imageDocWidth * imgRatio);

		r.addPicture(new FileInputStream(imageFile), imgFormat, imgFile, Units.toEMU(imageDocWidth), Units.toEMU(imageDocHeight));
	}

	private int getImageFormat(String imgFileName) {
		int format;
		if (imgFileName.endsWith(".emf"))
			format = XWPFDocument.PICTURE_TYPE_EMF;
		else if (imgFileName.endsWith(".wmf"))
			format = XWPFDocument.PICTURE_TYPE_WMF;
		else if (imgFileName.endsWith(".pict"))
			format = XWPFDocument.PICTURE_TYPE_PICT;
		else if (imgFileName.endsWith(".jpeg") || imgFileName.endsWith(".jpg"))
			format = XWPFDocument.PICTURE_TYPE_JPEG;
		else if (imgFileName.endsWith(".png"))
			format = XWPFDocument.PICTURE_TYPE_PNG;
		else if (imgFileName.endsWith(".dib"))
			format = XWPFDocument.PICTURE_TYPE_DIB;
		else if (imgFileName.endsWith(".gif"))
			format = XWPFDocument.PICTURE_TYPE_GIF;
		else if (imgFileName.endsWith(".tiff"))
			format = XWPFDocument.PICTURE_TYPE_TIFF;
		else if (imgFileName.endsWith(".eps"))
			format = XWPFDocument.PICTURE_TYPE_EPS;
		else if (imgFileName.endsWith(".bmp"))
			format = XWPFDocument.PICTURE_TYPE_BMP;
		else if (imgFileName.endsWith(".wpg"))
			format = XWPFDocument.PICTURE_TYPE_WPG;
		else {
			return 0;
		}
		return format;
	}

	public void close() throws IOException {
		stream = new FileOutputStream(filename);
		document.write(stream);
		stream.close();
		document.close();
	}

}
