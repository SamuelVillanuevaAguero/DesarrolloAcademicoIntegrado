/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package utilerias.busqueda;

import java.util.HashMap;
import java.util.Map;

/**
 *
 * @author Samue
 */
public class Docente {
    private int no;
    private String nombre;
    private Map<Integer, String> listaCursos;
    private int numeroCursosFD;
    private int numeroCursosAP;
    private String departamento;

    public Docente(int no, String nombre, int numeroCursosFD, int numeroCursosAP, String departamento) {
        this.no = no;
        this.nombre = nombre;
        this.listaCursos = new HashMap<>();
        this.numeroCursosFD = numeroCursosFD;
        this.numeroCursosAP = numeroCursosAP;
        this.departamento = departamento;
    }

    public int getNo() {
        return no;
    }

    public void setNo(int no) {
        this.no = no;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }

    public Map<Integer, String> getListaCursos() {
        return listaCursos;
    }

    public void setListaCursos(Map<Integer, String> listaCursos) {
        this.listaCursos = listaCursos;
    }

    public int getNumeroCursosFD() {
        return numeroCursosFD;
    }

    public void setNumeroCursosFD(int numeroCursosFD) {
        this.numeroCursosFD = numeroCursosFD;
    }

    public int getNumeroCursosAP() {
        return numeroCursosAP;
    }

    public void setNumeroCursosAP(int numeroCursosAP) {
        this.numeroCursosAP = numeroCursosAP;
    }

    public String getDepartamento() {
        return departamento;
    }

    public void setDepartamento(String departamento) {
        this.departamento = departamento;
    }
    
    @Override
    public String toString() {
        return nombre + "\n"
             + departamento + "\n"+
                numeroCursosAP + "\n"+
                numeroCursosFD;
    }
}
