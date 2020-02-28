package com.digiturtle.compiler;

import java.io.File;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileSystemView;

public class DocxImageCompiler {


	public static void main(String[] args) throws Exception {
		JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
		jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		int returnValue = jfc.showOpenDialog(null);
		if (returnValue == JFileChooser.APPROVE_OPTION) {
			File selectedFile = jfc.getSelectedFile();
			System.out.println(selectedFile.getAbsolutePath());
			
			DocxBuilder builder = new DocxBuilder("test1.docx");
			
			builder.addImageToWordDocument(new File("test/screen a 1.png"));
			builder.addImageToWordDocument(new File("test/screen a 2.png"));
			builder.addImageToWordDocument(new File("test/screen a 3.png"));
			builder.addImageToWordDocument(new File("test/screen a 4.png"));
			
			builder.close();			
		}
		else {
			throw new Exception("Directory required!");
		}
	}

	

}
