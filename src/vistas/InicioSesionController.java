/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/javafx/FXMLController.java to edit this template
 */
package vistas;

import java.net.URL;
import java.util.ResourceBundle;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.input.MouseEvent;
import utilerias.general.ControladorGeneral;

/**
 * FXML Controller class
 *
 * @author Samue
 */
public class InicioSesionController implements Initializable {

    /**
     * Initializes the controller class.
     */
    @FXML
    private Button botonCerrar;
    @FXML
    private Button botonMinimizar;
    
    //Métodos de los botones de la barra superior :)
    public void cerrarVentana() {
        Alert alerta = new Alert(Alert.AlertType.CONFIRMATION);
        alerta.setTitle("Sesión");
        alerta.setHeaderText(null);
        alerta.setContentText("¿Quieres salir del sistema?");

        if (alerta.showAndWait().get() == ButtonType.OK) {
            System.exit(0);
        }
    }
    
    public void minimizarVentana(MouseEvent event){
        ControladorGeneral.minimizarVentana(event);
    }

    @Override
    public void initialize(URL url, ResourceBundle rb) {
        // TODO
        botonCerrar.setOnMouseClicked(event -> {
            cerrarVentana();
        });
        
        botonMinimizar.setOnMouseClicked(event -> {
            minimizarVentana(event);
        });
    }
}
