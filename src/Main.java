import java.io.FileInputStream;
import java.io.IOException;

public class Product {
    private String name;
    private String category;
    private double price;
    private int stock;

    public Product(String name, String category, double price, int stock) {
        this.name = name;
        this.category = category;
        this.price = price;
        this.stock = stock;
    }

    // getters and setters


    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getCategory() {
        return category;
    }

    public void setCategory(String category) {
        this.category = category;
    }

    public double getPrice() {
        return price;
    }

    public void setPrice(double price) {
        this.price = price;
    }

    public int getStock() {
        return stock;
    }

    public void setStock(int stock) {
        this.stock = stock;
    }

    public static void main(String[] args) throws IOException {


        String fileName = "products.xlsx";
        FileInputStream fis = new FileInputStream(fileName);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        // read data from the "Products" sheet
        var sheet = workbook.getSheet("Products");
        var iterator = sheet.iterator();
        var products = new ArrayList<Product>();
        while (iterator.hasNext()) {
            var row = iterator.next();
            var nameCell = row.getCell(0);
            var categoryCell = row.getCell(1);
            var priceCell = row.getCell(2);
            var stockCell = row.getCell(3);

            // create product object
            var name = nameCell.getStringCellValue();
            var category = categoryCell.getStringCellValue();
            var price = priceCell.getNumericCellValue();
            var stock = (int) stockCell.getNumericCellValue();
            if (stock == 0) {
                stock = "OUT OF STOCK";
            }
            var product = new Product(name, category, price, stock);
            products.add(product);
        }

        // close workbook and file input stream
        workbook.close();
        fis.close();

        // do something with the products array
    }
}
