package example;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class DataBaseUtils {
    private static final String URL = "jdbc:postgresql://localhost:5432/Trial";
    private static final String USER = "postgres";
    private static final String PASSWORD = "5@Rosary";

    public static Connection getConnection() throws SQLException {
        return DriverManager.getConnection(URL, USER, PASSWORD);
    }
}
