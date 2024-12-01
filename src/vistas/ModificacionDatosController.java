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
    private ComboBox<String> departamento;

    private HashMap<String, Integer> docentesPorDepartamento = new HashMap<>();

    public void regresarVentana(MouseEvent event) throws IOException {
        // Verificar si hay datos en los campos
        boolean hayDatos = !directorField.getText().trim().isEmpty()
                || !coordinadorField.getText().trim().isEmpty()
                || !jefeDeptoField.getText().trim().isEmpty()
                || (totalDocentes.getValue() != null && totalDocentes.getValue() != 0)
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

    // Método para mostrar alertas en la aplicación
    private void mostrarAlerta(String titulo, String mensaje, AlertType tipo) {
        Alert alerta = new Alert(tipo);
        alerta.setTitle(titulo);
        alerta.setHeaderText(null);
        alerta.setContentText(mensaje);
        alerta.showAndWait();
    }

    //Métodos de los botones de la barra superior :)
    public void minimizarVentana(MouseEvent event) {
        ControladorGeneral.minimizarVentana(event);
    }

    public void restablecerCamposValoresIniciales() {
        // Solo si no hay registros previos
        if (noHayRegistrosPrevios()) {
            directorField.clear();
            coordinadorField.clear();
            jefeDeptoField.clear();

            for (String depto : departamento.getItems()) {
                docentesPorDepartamento.put(depto, 0);
            }

            totalDocentes.getValueFactory().setValue(0);
            año.getValueFactory().setValue(Calendar.getInstance().get(Calendar.YEAR));
            periodoEscolar.getSelectionModel().clearSelection();
            departamento.getSelectionModel().clearSelection();
        } else {
            mostrarAlerta("Aviso", "No se pueden borrar los datos existentes. Utilice el botón Guardar para actualizar la información.", AlertType.WARNING);
        }
    }

    private boolean noHayRegistrosPrevios() {
        String rutaArchivo = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\informacion_modificable\\info.xlsx";
        File archivo = new File(rutaArchivo);

        if (!archivo.exists()) {
            return true;
        }

        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(archivo))) {
            Sheet sheet = workbook.getSheet(String.valueOf(año.getValue()));
            return sheet == null || sheet.getPhysicalNumberOfRows() == 0;
        } catch (IOException e) {
            return true;
        }
    }

    public void guardarDatosEnExcel(MouseEvent event) {
        if (!validarCampos()) {
            return;
        }

        String rutaArchivo = ControladorGeneral.obtenerRutaDeEjecusion() + "\\Gestion_de_Cursos\\Sistema\\informacion_modificable\\info.xlsx";
        File archivo = new File(rutaArchivo);

        try {
            Workbook workbook;
            Sheet sheet;

            if (archivo.exists()) {
                workbook = new XSSFWorkbook(new FileInputStream(archivo));
                sheet = workbook.getSheet(String.valueOf(año.getValue()));

                if (sheet == null) {
                    sheet = workbook.createSheet(String.valueOf(año.getValue()));
                    configurarCabecerasHoja(sheet);
                } else {
                    // Cargar datos existentes en la vista
                    cargarDatosExistentes(sheet);
                }
            } else {
                workbook = new XSSFWorkbook();
                sheet = workbook.createSheet(String.valueOf(año.getValue()));
                configurarCabecerasHoja(sheet);
            }

            // Actualizar datos en la hoja
            actualizarDatosEnHoja(sheet);
            
            // Guardar archivo
            try (FileOutputStream fileOut = new FileOutputStream(rutaArchivo)) {
                workbook.write(fileOut);
                mostrarAlerta("Éxito", "Datos guardados exitosamente", AlertType.INFORMATION);
                restablecerCamposValoresIniciales();
            }
        } catch (IOException e) {
            mostrarAlerta("Error", "No se pudo guardar el archivo: " + e.getMessage(), AlertType.ERROR);
            e.printStackTrace();
        }
    }

    private void cargarDatosExistentes(Sheet sheet) {
        // Obtener el periodo seleccionado
        String periodoSeleccionado = periodoEscolar.getValue();

        // Configurar rangos de filas según el periodo
        int filaInicial, filaTotal;
        if (periodoSeleccionado.equals("Enero-Julio")) {
            filaInicial = 4;
            filaTotal = 12;
        } else { // Agosto-Diciembre
            filaInicial = 13;
            filaTotal = 21;
        }

        try {
            // Cargar datos de las primeras filas (siempre iguales)
            Row directorRow = sheet.getRow(0);
            if (directorRow != null && directorRow.getCell(1) != null) {
                directorField.setText(directorRow.getCell(1).getStringCellValue());
            }

            Row jefeDeptoRow = sheet.getRow(1);
            if (jefeDeptoRow != null && jefeDeptoRow.getCell(1) != null) {
                jefeDeptoField.setText(jefeDeptoRow.getCell(1).getStringCellValue());
            }

            Row coordinadorRow = sheet.getRow(2);
            if (coordinadorRow != null && coordinadorRow.getCell(1) != null) {
                coordinadorField.setText(coordinadorRow.getCell(1).getStringCellValue());
            }

            // Limpiar mapa de docentes
            docentesPorDepartamento.clear();
            int totalDocentesValor = 0;

            // Cargar datos de departamentos para el periodo seleccionado
            for (int i = 0; i < departamento.getItems().size(); i++) {
                String depto = departamento.getItems().get(i);
                Row row = sheet.getRow(filaInicial + i);

                if (row != null && row.getCell(1) != null) {
                    int docentes = (int) row.getCell(1).getNumericCellValue();
                    docentesPorDepartamento.put(depto, docentes);
                    totalDocentesValor += docentes;
                } else {
                    docentesPorDepartamento.put(depto, 0);
                }
            }

            // Establecer valores
            totalDocentes.getValueFactory().setValue(totalDocentesValor);

        } catch (Exception e) {
            mostrarAlerta("Error", "No se pudieron cargar los datos existentes: " + e.getMessage(), AlertType.ERROR);
            e.printStackTrace();
        }
    }

    private void actualizarDatosEnHoja(Sheet sheet) {
        // Obtener el periodo seleccionado
        String periodoSeleccionado = periodoEscolar.getValue();

        // Configurar rangos de filas según el periodo
        int filaInicial;
        if (periodoSeleccionado.equals("Enero-Julio")) {
            filaInicial = 4;
        } else { // Agosto-Diciembre
            filaInicial = 13;
        }

        // Actualizar datos generales
        Row directorRow = sheet.getRow(0);
        if (directorRow.getCell(1) == null) {
            directorRow.createCell(1);
        }
        directorRow.getCell(1).setCellValue(directorField.getText());

        Row jefeDeptoRow = sheet.getRow(1);
        if (jefeDeptoRow.getCell(1) == null) {
            jefeDeptoRow.createCell(1);
        }
        jefeDeptoRow.getCell(1).setCellValue(jefeDeptoField.getText());

        Row coordinadorRow = sheet.getRow(2);
        if (coordinadorRow.getCell(1) == null) {
            coordinadorRow.createCell(1);
        }
        coordinadorRow.getCell(1).setCellValue(coordinadorField.getText());

        // Actualizar periodo
        Row periodoRow = sheet.getRow(3);
        if (periodoRow.getCell(1) == null) {
            periodoRow.createCell(1);
        }
        periodoRow.getCell(1).setCellValue(periodoSeleccionado);

        // Calcular suma total de docentes
        int totalDocentes = 0;
        for (String depto : departamento.getItems()) {
            totalDocentes += docentesPorDepartamento.get(depto);
        }

        // Crear o actualizar la celda de suma total de docentes
        Row sumaDocentesRow = sheet.getRow(filaInicial - 1);
        if (sumaDocentesRow == null) {
            sumaDocentesRow = sheet.createRow(filaInicial - 1);
        }
        if (sumaDocentesRow.getCell(1) == null) {
            sumaDocentesRow.createCell(1);
        }
        sumaDocentesRow.getCell(1).setCellValue(totalDocentes);

        // Actualizar datos de departamentos
        for (int i = 0; i < departamento.getItems().size(); i++) {
            String depto = departamento.getItems().get(i);
            Row row = sheet.getRow(filaInicial + i);

            if (row.getCell(1) == null) {
                row.createCell(1);
            }

            row.getCell(1).setCellValue(docentesPorDepartamento.get(depto));
        }
    }

    private void configurarCabecerasHoja(Sheet sheet) {
        // Establecer cabeceras fijas en las primeras filas
        sheet.createRow(0).createCell(0).setCellValue("Director:");
        sheet.createRow(1).createCell(0).setCellValue("Jefe de Departamento:");
        sheet.createRow(2).createCell(0).setCellValue("Coordinador:");
        sheet.createRow(3).createCell(0).setCellValue("Periodo:");

        // Crear filas para departamentos
        String[] departamentos = {
            "CIENCIAS BASICAS",
            "CIENCIAS ECONOMICO ADMINISTRATIVO",
            "CIENCIAS DE LA TIERRA",
            "INGENIERIA INDUSTRIAL",
            "METAL MECANICA",
            "QUIMICA Y BIOQUIMICA",
            "SISTEMAS COMPUTACIONALES",
            "POSGRADO"
        };

        // Para Enero-Julio
        for (int i = 0; i < departamentos.length; i++) {
            Row row = sheet.getRow(i + 4) != null ? sheet.getRow(i + 4) : sheet.createRow(i + 4);
            row.createCell(0).setCellValue(departamentos[i]);
        }

        // Para Agosto-Diciembre
        for (int i = 0; i < departamentos.length; i++) {
            Row row = sheet.getRow(i + 13) != null ? sheet.getRow(i + 13) : sheet.createRow(i + 13);
            row.createCell(0).setCellValue(departamentos[i]);
        }
    }

    private boolean validarCampos() {
        // Campos obligatorios siempre
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

        // Validar que cada departamento tenga un valor entre 10 y 80
        boolean departamentosValidos = true;
        StringBuilder mensaje = new StringBuilder("Los siguientes departamentos requieren un valor entre 10 y 80:\n");

        for (String depto : departamento.getItems()) {
            int valorDocentes = docentesPorDepartamento.get(depto);
            if (valorDocentes < 10 || valorDocentes > 80) {
                mensaje.append("- ").append(depto).append("\n");
                departamentosValidos = false;
            }
        }

        if (!departamentosValidos) {
            mostrarAlerta("Validación", mensaje.toString(), AlertType.WARNING);
            return false;
        }

        // Validaciones de formato de nombre
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

        return true;
    }

    @Override
    public void initialize(URL url, ResourceBundle rb) {

        departamento.getItems().addAll(
                "CIENCIAS BASICAS",
                "CIENCIAS ECONOMICO ADMINISTRATIVO",
                "CIENCIAS DE LA TIERRA",
                "INGENIERIA INDUSTRIAL",
                "METAL MECANICA",
                "QUIMICA Y BIOQUIMICA",
                "SISTEMAS COMPUTACIONALES",
                "POSGRADO"
        );

        // Inicializar valores por defecto en el HashMap
        for (String depto : departamento.getItems()) {
            docentesPorDepartamento.put(depto, 0); // Inicializar en 0
        }

        // Configuración del Spinner de docentes
        SpinnerValueFactory<Integer> docentesFactory = new SpinnerValueFactory.IntegerSpinnerValueFactory(0, 80, 0);
        totalDocentes.setValueFactory(docentesFactory);
        totalDocentes.setEditable(true);

        // Configuración del Spinner de año
        int añoActual = Calendar.getInstance().get(Calendar.YEAR);
        SpinnerValueFactory<Integer> añoFactory = new SpinnerValueFactory.IntegerSpinnerValueFactory(2000, 2100, añoActual);
        año.setValueFactory(añoFactory);
        año.setEditable(true);

        // Actualizar el Spinner de docentes al seleccionar un departamento
        departamento.getSelectionModel().selectedItemProperty().addListener((observable, oldValue, newValue) -> {
            if (newValue != null) {
                int valorDocentes = docentesPorDepartamento.getOrDefault(newValue, 0);
                totalDocentes.getValueFactory().setValue(valorDocentes);
            }
        });

        // Guardar cambios en el HashMap cuando se edita el Spinner
        totalDocentes.valueProperty().addListener((observable, oldValue, newValue) -> {
            String deptoSeleccionado = departamento.getValue();
            if (deptoSeleccionado != null) {
                docentesPorDepartamento.put(deptoSeleccionado, newValue);
            }
        });

        periodoEscolar.getItems().addAll("Enero-Julio", "Agosto-Diciembre");

        // Configurar eventos de botones
        botonGuardar.setOnMouseClicked(this::guardarDatosEnExcel);
        botonCerrar.setOnMouseClicked(event -> {
            try {
                ControladorGeneral.cerrarVentana(event, "¿Quieres cerrar sesión??", getClass());
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

        // El botón de limpiar no está disponible cuando hay registros previos
        botonLimpiar.setDisable(!noHayRegistrosPrevios());
        botonLimpiar.setOnMouseClicked(event -> restablecerCamposValoresIniciales());
    }
}
