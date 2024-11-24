/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/javafx/FXMLController.java to edit this template
 */
package vistas;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.net.URL;
import java.util.ResourceBundle;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.input.MouseEvent;
import javafx.stage.Stage;
import utilerias.general.ControladorGeneral;
import java.util.Optional;

public class ModificacionDatosController implements Initializable {

    @FXML
    private Button botonCerrar;
    @FXML
    private Button botonMinimizar;
    @FXML
    private Button botonRegresar;
    @FXML
    private Button botonGuardar;
    @FXML
    private Button botonLimpiar;
    @FXML
    private TextField directorField;
    @FXML
    private TextField coordinadorField;
    @FXML
    private TextField jefeDeptoField;
    @FXML
    private Spinner<Integer> totalDocentesField;
    @FXML
    private Spinner<Integer> año;
    @FXML
    private ComboBox<String> periodoEscolar;

    public void guardarDatosEnExcel(MouseEvent event) {
        if (!validarCampos()) {
            return;
        }

        String rutaArchivo = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\informacion_modificable\\info.xlsx";
        File archivo = new File(rutaArchivo);
        Workbook workbook;

        try {
            // Verificar si el archivo existe
            if (archivo.exists()) {
                // Si existe, cargar el archivo existente
                workbook = new XSSFWorkbook(new FileInputStream(archivo));
            } else {
                // Si no existe, crear uno nuevo
                workbook = new XSSFWorkbook();
            }

            // Obtener o crear la hoja para el año seleccionado
            String nombreHoja = String.valueOf(año.getValue());
            Sheet sheet = workbook.getSheet(nombreHoja);
            if (sheet == null) {
                sheet = workbook.createSheet(nombreHoja);
            }

            // Configurar las cabeceras si es una hoja nueva
            if (sheet.getPhysicalNumberOfRows() == 0) {
                Row headerRow = sheet.createRow(0);
                headerRow.createCell(0).setCellValue("Director:");
                headerRow = sheet.createRow(1);
                headerRow.createCell(0).setCellValue("Jefe de Departamento:");
                headerRow = sheet.createRow(2);
                headerRow.createCell(0).setCellValue("Coordinador:");
                headerRow = sheet.createRow(3);
                headerRow.createCell(0).setCellValue("Periodo Enero-Julio:");
                headerRow = sheet.createRow(4);
                headerRow.createCell(0).setCellValue("Periodo Agosto-Diciembre:");
            }

            // Actualizar datos básicos
            sheet.getRow(0).createCell(1).setCellValue(directorField.getText());
            sheet.getRow(1).createCell(1).setCellValue(jefeDeptoField.getText());
            sheet.getRow(2).createCell(1).setCellValue(coordinadorField.getText());

            // Actualizar período y total de docentes
            int periodo = periodoEscolar.getValue().equals("Enero-Julio") ? 1 : 2;
            int fila = periodo == 1 ? 3 : 4; // Fila 4 para Enero-Julio, Fila 5 para Agosto-Diciembre
            
            Row row = sheet.getRow(fila);
            row.createCell(1).setCellValue(totalDocentesField.getValue());
            //row.createCell(2).setCellValue(totalDocentesField.getValue());

            // Ajustar el ancho de las columnas automáticamente
            for (int i = 0; i < 3; i++) {
                sheet.autoSizeColumn(i);
            }

            // Guardar archivo
            try (FileOutputStream fileOut = new FileOutputStream(rutaArchivo)) {
                workbook.write(fileOut);
                mostrarAlerta("Éxito", "Datos guardados exitosamente", AlertType.INFORMATION);
            }

        } catch (IOException e) {
            mostrarAlerta("Error", "No se pudo guardar el archivo: " + e.getMessage(), AlertType.ERROR);
            e.printStackTrace();
        }
    }

    private boolean validarCampos() {
        if (!directorField.getText().matches("[a-zA-ZáéíóúÁÉÍÓÚ\\s]+")) {
            mostrarAlerta("Validación", "El campo 'Director' solo debe contener letras.", AlertType.WARNING);
            return false;
        }
        if (!coordinadorField.getText().matches("[a-zA-ZZáéíóúÁÉÍÓÚ\\s]+")) {
            mostrarAlerta("Validación", "El campo 'Coordinador' solo debe contener letras.", AlertType.WARNING);
            return false;
        }
        if (!jefeDeptoField.getText().matches("[a-zA-ZZáéíóúÁÉÍÓÚ\\s]+")) {
            mostrarAlerta("Validación", "El campo 'Jefe de Departamento' solo debe contener letras.", AlertType.WARNING);
            return false;
        }
        if (totalDocentesField.getValue() == null || totalDocentesField.getValue() < 0) {
            mostrarAlerta("Validación", "El campo 'Total de Docentes' solo debe contener números positivos.", AlertType.WARNING);
            return false;
        }
        if (periodoEscolar.getValue() == null) {
            mostrarAlerta("Validación", "Por favor selecciona un período escolar.", AlertType.WARNING);
            return false;
        }
        return true;
    }

    @Override
    public void initialize(URL url, ResourceBundle rb) {
        SpinnerValueFactory<Integer> valueFactory = new SpinnerValueFactory.IntegerSpinnerValueFactory(1, 800, 100);
        totalDocentesField.setValueFactory(valueFactory);
        totalDocentesField.setEditable(true);

        SpinnerValueFactory<Integer> añoValueFactory = new SpinnerValueFactory.IntegerSpinnerValueFactory(2000, 2100, 2024);
        año.setValueFactory(añoValueFactory);

        periodoEscolar.getItems().addAll("Enero-Julio", "Agosto-Diciembre");

        // Configurar eventos de botones
        botonGuardar.setOnMouseClicked(this::guardarDatosEnExcel);
        botonCerrar.setOnMouseClicked(this::cerrarVentana);
        botonMinimizar.setOnMouseClicked(this::minimizarVentana);
        botonRegresar.setOnMouseClicked(event -> {
            try {
                regresarVentana(event);
            } catch (IOException ex) {
                Logger.getLogger(ModificacionDatosController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });
        botonLimpiar.setOnMouseClicked(event -> limpiarCampos());
    }

    // Los demás métodos (cerrarVentana, minimizarVentana, regresarVentana, limpiarCampos, mostrarAlerta) 
    // permanecen igual que en tu código original


    //Métodos de los botones de la barra superior :)
    public void cerrarVentana(MouseEvent event) {
        Alert confirmacion = new Alert(Alert.AlertType.CONFIRMATION);
        confirmacion.setTitle("Confirmación");
        confirmacion.setHeaderText(null);
        confirmacion.setContentText("¿Quieres cerrar sesión?");

        Optional<ButtonType> resultado = confirmacion.showAndWait();
        if (resultado.isPresent() && resultado.get() == ButtonType.OK) {
            Stage stage = (Stage) ((Button) event.getSource()).getScene().getWindow();
            stage.close();
        }
    }

    public void minimizarVentana(MouseEvent event) {
        ControladorGeneral.minimizarVentana(event);
    }

    public void regresarVentana(MouseEvent event) throws IOException {
        // Verificar si hay datos en los campos
        boolean hayDatos = !directorField.getText().isEmpty() || !coordinadorField.getText().isEmpty()
                || !jefeDeptoField.getText().isEmpty() || totalDocentesField.getValue() != null;

        // Si hay datos, mostrar mensaje de confirmación para guardar
        if (hayDatos) {
            Alert confirmacion = new Alert(Alert.AlertType.CONFIRMATION);
            confirmacion.setTitle("Confirmación");
            confirmacion.setHeaderText("¿Deseas guardar los datos?");
            confirmacion.setContentText("Tienes datos ingresados en los campos. ¿Deseas guardarlos antes de regresar?");

            ButtonType botonGuardar = new ButtonType("Guardar");
            ButtonType botonNoGuardar = new ButtonType("No Guardar");
            ButtonType botonCancelar = new ButtonType("Cancelar", ButtonBar.ButtonData.CANCEL_CLOSE);

            confirmacion.getButtonTypes().setAll(botonGuardar, botonNoGuardar, botonCancelar);
            Optional<ButtonType> resultado = confirmacion.showAndWait();

            if (resultado.isPresent()) {
                if (resultado.get() == botonGuardar) {
                    guardarDatosEnExcel(event); // Llamar a método de guardar
                    ControladorGeneral.regresar(event, "Principal", getClass()); // Regresar a la ventana principal
                } else if (resultado.get() == botonNoGuardar) {
                    ControladorGeneral.regresar(event, "Principal", getClass()); // Regresar sin guardar
                }
                // Si selecciona cancelar, no se realiza ninguna acción adicional
            }
        } else {
            // Si no hay datos, regresar sin mostrar mensaje
            ControladorGeneral.regresar(event, "Principal", getClass());
        }
    }

    public void limpiarCampos() {
        directorField.clear();
        coordinadorField.clear();
        jefeDeptoField.clear();
        totalDocentesField.getValueFactory().setValue(0); // Restablecer Spinner
        año.getValueFactory().setValue(2024); // Restablecer Spinner de año al valor predeterminado
        periodoEscolar.getSelectionModel().clearSelection(); // Limpiar selección del ComboBox

    }

   
    // Método para mostrar alertas en la aplicación
    private void mostrarAlerta(String titulo, String mensaje, AlertType tipo) {
        Alert alerta = new Alert(tipo);
        alerta.setTitle(titulo);
        alerta.setHeaderText(null);
        alerta.setContentText(mensaje);
        alerta.showAndWait();
    }

    // Método para validar los campos de entrada
    /*private boolean validarCampos() {
        // Validar que los campos de nombres solo contengan letras
        if (!directorField.getText().matches("[a-zA-Z\\s]+")) {
            mostrarAlerta("Validación", "El campo 'Director' solo debe contener letras.", AlertType.WARNING);
            return false;
        }
        if (!coordinadorField.getText().matches("[a-zA-Z\\s]+")) {
            mostrarAlerta("Validación", "El campo 'Coordinador' solo debe contener letras.", AlertType.WARNING);
            return false;
        }
        if (!jefeDeptoField.getText().matches("[a-zA-Z\\s]+")) {
            mostrarAlerta("Validación", "El campo 'Jefe de Departamento' solo debe contener letras.", AlertType.WARNING);
            return false;
        }

        // Validar que el campo de total de docentes solo contenga números
        if (totalDocentesField.getValue() == null || totalDocentesField.getValue() < 0) {
            mostrarAlerta("Validación", "El campo 'Total de Docentes' solo debe contener números positivos.", AlertType.WARNING);
            return false;
        }
        if (periodoEscolar.getValue() == null) {
            mostrarAlerta("Validación", "Por favor selecciona un período escolar.", AlertType.WARNING);
            return false;
        }

        return true;
    }*/

   /* @Override
    public void initialize(URL url, ResourceBundle rb) {
        // Otros botones

        // Configuración del Spinner para aceptar valores enteros
        SpinnerValueFactory<Integer> valueFactory = new SpinnerValueFactory.IntegerSpinnerValueFactory(100, 800); // Rango de 0 a 100, ajusta según necesites
        totalDocentesField.setValueFactory(valueFactory);
        // Configuración del botón Guardar para que llame al método guardarDatosEnExcel
        botonGuardar.setOnMouseClicked(event -> guardarDatosEnExcel(event));

        botonCerrar.setOnMouseClicked(event -> {
            cerrarVentana(event);
        });

        botonMinimizar.setOnMouseClicked(this::minimizarVentana);
        botonRegresar.setOnMouseClicked(event -> {
            try {
                regresarVentana(event);
            } catch (IOException ex) {
                Logger.getLogger(BusquedaEstadisticaController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });
        botonLimpiar.setOnMouseClicked(event -> limpiarCampos());

        // Configuración del Spinner para el año
        SpinnerValueFactory<Integer> añoValueFactory = new SpinnerValueFactory.IntegerSpinnerValueFactory(2000, 2100, 2024);
        año.setValueFactory(añoValueFactory);

        // Configuración del ComboBox para el período escolar
        periodoEscolar.getItems().addAll("Enero-Julio", "Agosto-Diciembre");
    }*/
}
