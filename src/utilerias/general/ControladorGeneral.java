/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package utilerias.general;

import java.io.IOException;
import javafx.event.ActionEvent;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.input.MouseEvent;
import javafx.stage.Stage;
import vistas.PrincipalController;

/**
 *
 * @author Samue
 */
public class ControladorGeneral {

    public static void cerrarVentana(MouseEvent event, String mensajeConfirmacion, Class clase) throws IOException {
        Alert alerta = new Alert(Alert.AlertType.CONFIRMATION);
        alerta.setTitle("Sesión");
        alerta.setHeaderText(null);
        alerta.setContentText(mensajeConfirmacion);

        if (alerta.showAndWait().get() == ButtonType.OK) {
            Parent root = FXMLLoader.load(clase.getResource("/vistas/InicioSesion.fxml"));
            Stage stage = (Stage) ((javafx.scene.Node) event.getSource()).getScene().getWindow();

            Scene scene = new Scene(root);
            stage.setScene(scene);
            stage.show();
        }

    }

    public static void minimizarVentana(MouseEvent event) {
        ((Stage) ((Button)event.getSource()).getScene().getWindow()).setIconified(true);
    }

    public static void regresar(MouseEvent event, String nombreVista, Class clase) throws IOException {
        FXMLLoader loader = new FXMLLoader(clase.getResource("/vistas/" + nombreVista + ".fxml"));
        Parent root = loader.load();

        Stage stage = (Stage) ((javafx.scene.Node) event.getSource()).getScene().getWindow();
        Scene scene = new Scene(root);
        stage.setScene(scene);
        stage.show();
    }
    
}
