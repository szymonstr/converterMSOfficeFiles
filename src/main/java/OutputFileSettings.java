public class OutputFileSettings {

    private boolean pdf;
    private boolean png;
    private int[] numPages;
    private String path;
    private String fileName;

    public OutputFileSettings(boolean pdf, boolean png, int[] numPages, String path, String fileName) {
        this.pdf = pdf;
        this.png = png;
        this.numPages = numPages;
        this.path = path;
        this.fileName = fileName;
    }

    public boolean isPdf() {
        return pdf;
    }

    public boolean isPng() {
        return png;
    }

    public int[] getNumPages() {
        return numPages;
    }

    public String getPath() {
        return path;
    }

    public String getFileName() {
        return fileName;
    }
}
