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
import java.time.Year;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Map;
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
    private Spinner<Integer> totalDocentes;
    @FXML
    private Spinner<Integer> año;
    @FXML
    private ComboBox<String> periodoEscolar;
    @FXML
    private ComboBox<String> departamentos;

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
            row.createCell(1).setCellValue(totalDocentes.getValue());

            // Ajustar el ancho de las columnas automáticamente
            for (int i = 0; i < 3; i++) {
                sheet.autoSizeColumn(i);
            }

            // Guardar archivo
            try (FileOutputStream fileOut = new FileOutputStream(rutaArchivo)) {
                workbook.write(fileOut);
                mostrarAlerta("Éxito", "Datos guardados exitosamente", AlertType.INFORMATION);

                // Limpiar campos después de guardar exitosamente
                restablecerCamposValoresIniciales();
            }
        } catch (IOException e) {
            mostrarAlerta("Error", "No se pudo guardar el archivo: " + e.getMessage(), AlertType.ERROR);
            e.printStackTrace();
        }
    }

    @Override
    public void initialize(URL url, ResourceBundle rb) {
        // Configuración del Spinner de docentes
        SpinnerValueFactory<Integer> docentesFactory = new SpinnerValueFactory.IntegerSpinnerValueFactory(1, 100, 10);
        totalDocentes.setValueFactory(docentesFactory);
        totalDocentes.setEditable(true);

        // Configuración del Spinner de año
        int añoActual = Calendar.getInstance().get(Calendar.YEAR); // Año actual fijo para el sistema
        SpinnerValueFactory<Integer> añoFactory = new SpinnerValueFactory.IntegerSpinnerValueFactory(2000, 2100, añoActual);
        año.setValueFactory(añoFactory);
        año.setEditable(true);

        // Validación para el Spinner de docentes
        totalDocentes.getEditor().focusedProperty().addListener((observable, oldValue, newValue) -> {
            if (!newValue) { // Cuando pierde el foco
                try {
                    String text = totalDocentes.getEditor().getText();
                    if (text.isEmpty()) {
                        totalDocentes.getValueFactory().setValue(10);
                    } else {
                        int value = Integer.parseInt(text);
                        if (value < 1) {
                            mostrarAlerta("Validación", "El número mínimo de docentes debe ser 100", AlertType.WARNING);
                            totalDocentes.getValueFactory().setValue(10);
                        }
                    }
                } catch (NumberFormatException e) {
                    totalDocentes.getValueFactory().setValue(10);
                }
            }
        });

        // Validación corregida para el Spinner de año
        año.getEditor().focusedProperty().addListener((observable, oldValue, newValue) -> {
            if (!newValue) { // Cuando pierde el foco
                try {
                    String text = año.getEditor().getText();
                    int value;

                    if (text.isEmpty()) {
                        año.getValueFactory().setValue(añoActual);
                    } else {
                        value = Integer.parseInt(text);
                        if (value < 2000 || value > 2100) {
                            mostrarAlerta("Validación", "El año debe estar entre 2000 y 2100", AlertType.WARNING);
                            año.getValueFactory().setValue(añoActual);

                        } else {
                            año.getValueFactory().setValue(value);
                        }
                    }
                } catch (NumberFormatException e) {
                    año.getValueFactory().setValue(añoActual);
                }
            }
        });

        // Solo permitir números en los Spinners
        totalDocentes.getEditor().textProperty().addListener((observable, oldValue, newValue) -> {
            if (!newValue.matches("\\d*")) {
                totalDocentes.getEditor().setText(oldValue);
            }
        });

        // Solo permitir números en el Spinner de año
        año.getEditor().textProperty().addListener((observable, oldValue, newValue) -> {
            if (!newValue.matches("\\d*")) {
                año.getEditor().setText(oldValue);
            } else if (!newValue.isEmpty()) {
                try {
                    int value = Integer.parseInt(newValue);
                    if (value > 2100) {
                        año.getEditor().setText(oldValue);
                    }
                } catch (NumberFormatException e) {
                    año.getEditor().setText(oldValue);
                }
            }
        });

        periodoEscolar.getItems().addAll("Enero-Julio", "Agosto-Diciembre");

        // Configurar eventos de botones
        botonGuardar.setOnMouseClicked(this::guardarDatosEnExcel);
        botonCerrar.setOnMouseClicked(event -> {
            try {
                ControladorGeneral.cerrarVentana(event, "¿Quieres cerrar sesión??", getClass()
                );
            } catch (IOException ex) {
                Logger.getLogger(ModificacionDatosController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });

        botonMinimizar.setOnMouseClicked(this::minimizarVentana);
        botonRegresar.setOnMouseClicked(event -> {
            try {
                regresarVentana(event);
            } catch (IOException ex) {
                Logger.getLogger(ModificacionDatosController.class.getName()).log(Level.SEVERE, null, ex);
            }
        });
        botonLimpiar.setOnMouseClicked(event -> restablecerCamposValoresIniciales());

        //Departamentos
        Map<String, Integer> mapaDepartamentos = new HashMap<>();
        departamentos.getItems().addAll("CIENCIAS BÁSICAS", "CIENCIAS ECONÓMICO ADMINISTRATIVAS", "CIENCIAS DE LA TIERRA", "INGENIERÍA INDUSTRIAL", "METAL MECÁNICA", "QUÍMICA Y BIOQUÍMICA", "SISTEMAS COMPUTACIONALES", "POSGRADO");

        totalDocentes.getEditor().focusedProperty().addListener((observable, oldValue, newValue) -> {
            String departamento = departamentos.getValue();
            int docentes = totalDocentes.getValue();

            if (mapaDepartamentos.containsKey(departamento)) {
                mapaDepartamentos.remove(departamento);
                mapaDepartamentos.put(departamento, docentes);
            } else {
                mapaDepartamentos.put(departamento, docentes);
            }
            System.out.println(mapaDepartamentos);
        });

        departamentos.setOnAction(event -> {
            String departamento = departamentos.getValue();

            if (mapaDepartamentos.containsKey(departamento)) {
                totalDocentes.getValueFactory().setValue(mapaDepartamentos.get(departamento));
            } else {
                totalDocentes.getValueFactory().setValue(10);
            }
        });
    }


    private boolean validarCampos() {
        // Verificar si todos los campos están vacíos
        boolean todosVacios = directorField.getText().trim().isEmpty()
                && coordinadorField.getText().trim().isEmpty()
                && jefeDeptoField.getText().trim().isEmpty()
                && periodoEscolar.getValue() == null;

        if (todosVacios) {
            mostrarAlerta("Validación", "Todos los campos son obligatorios. Por favor, complete el formulario.", AlertType.WARNING);
            return false;
        }

        // Validaciones individuales
        if (directorField.getText().trim().isEmpty()) {
            mostrarAlerta("Validación", "Por favor, ingrese el nombre del Director.", AlertType.WARNING);
            directorField.requestFocus();
            return false;
        }

        if (coordinadorField.getText().trim().isEmpty()) {
            mostrarAlerta("Validación", "Por favor, ingrese el nombre del Coordinador.", AlertType.WARNING);
            coordinadorField.requestFocus();
            return false;
        }

        if (jefeDeptoField.getText().trim().isEmpty()) {
            mostrarAlerta("Validación", "Por favor, ingrese el nombre del Jefe de Departamento.", AlertType.WARNING);
            jefeDeptoField.requestFocus();
            return false;
        }

        if (periodoEscolar.getValue() == null) {
            mostrarAlerta("Validación", "Por favor, seleccione un periodo escolar.", AlertType.WARNING);
            periodoEscolar.requestFocus();
            return false;
        }

        // Validar formato de nombres solo si no están vacíos
        if (!directorField.getText().matches("[a-zA-ZáéíóúÁÉÍÓÚ\\s]+")) {
            mostrarAlerta("Validación", "El campo 'Director' solo debe contener letras.", AlertType.WARNING);
            directorField.requestFocus();
            return false;
        }
        if (!coordinadorField.getText().matches("[a-zA-ZáéíóúÁÉÍÓÚ\\s]+")) {
            mostrarAlerta("Validación", "El campo 'Coordinador' solo debe contener letras.", AlertType.WARNING);
            coordinadorField.requestFocus();
            return false;
        }
        if (!jefeDeptoField.getText().matches("[a-zA-ZáéíóúÁÉÍÓÚ\\s]+")) {
            mostrarAlerta("Validación", "El campo 'Jefe de Departamento' solo debe contener letras.", AlertType.WARNING);
            jefeDeptoField.requestFocus();
            return false;
        }

        // Validar número de docentes
        try {
            int docentes = Integer.parseInt(totalDocentes.getEditor().getText());
            if (docentes < 100) {
                mostrarAlerta("Validación", "El número mínimo de docentes debe ser 100.", AlertType.WARNING);
                totalDocentes.requestFocus();
                return false;
            }
        } catch (NumberFormatException e) {
            mostrarAlerta("Validación", "Por favor ingrese un número válido de docentes.", AlertType.WARNING);
            totalDocentes.requestFocus();
            return false;
        }

        return true;
    }
    // Los demás métodos (cerrarVentana, minimizarVentana, regresarVentana, limpiarCampos, mostrarAlerta) 
    // permanecen igual que en tu código original

    //Métodos de los botones de la barra superior :)
    public void minimizarVentana(MouseEvent event) {
        ControladorGeneral.minimizarVentana(event);
    }

    public void regresarVentana(MouseEvent event) throws IOException {
        // Verificar si hay datos en los campos
        boolean hayDatos = !directorField.getText().trim().isEmpty()
                || !coordinadorField.getText().trim().isEmpty()
                || !jefeDeptoField.getText().trim().isEmpty()
                || (totalDocentes.getValue() != null && totalDocentes.getValue() != 100)
                || periodoEscolar.getValue() != null;
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
                    guardarDatosEnExcel(event);
                    ControladorGeneral.regresar(event, "Principal", getClass());
                } else if (resultado.get() == botonNoGuardar) {
                    ControladorGeneral.regresar(event, "Principal", getClass());
                }
            }
        } else {
            ControladorGeneral.regresar(event, "Principal", getClass());
        }
    }

    // Nuevo método para restablecer campos a valores iniciales
    public void restablecerCamposValoresIniciales() {
        directorField.clear();
        coordinadorField.clear();
        jefeDeptoField.clear();
        totalDocentes.getValueFactory().setValue(100); // Valor inicial del Spinner
        año.getValueFactory().setValue(Calendar.getInstance().get(Calendar.YEAR)); // Valor inicial del año
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

}
