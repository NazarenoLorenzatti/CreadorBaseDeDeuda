package FileReader;

import FileWriter.EscritorTxt;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import static java.lang.String.valueOf;
import java.math.BigDecimal;
import java.util.Scanner;
import javax.swing.JTextArea;
import javax.swing.JTextField;

public class LectorTXT_DESUSO {

    private String path;
    private String nuevaLinea;
    private EscritorTxt escritor;
    private double total = 0;
    

    public LectorTXT_DESUSO(String path) {
        this.path = path;
        path = path.replaceAll(".txt", " - procesado.txt");
        escritor = new EscritorTxt(path);
    }

    public void leerArchivo(JTextArea area, JTextField jText) throws FileNotFoundException, IOException {

        System.out.println("INICIO DEL PROCESADO POR FAVOR AGUARDE");
        
        int j = 0;
        int contador = 0;        

        InputStream ins = new FileInputStream(this.path);
        InputStream ins2 = new FileInputStream(this.path);

        Scanner scanner = new Scanner(ins);
        Scanner scanner2 = new Scanner(ins2);

        // creamos un segundo scanner para contar el total de filas del archivo .TXT
        while (scanner2.hasNextLine()) {
            contador++;
            scanner2.nextLine();
        }
        scanner2.close();

        // Hago un array de los Dni para poder comparar si el cliente tiene mas de dos facturas
        String[] dniClientes = new String[contador];

        // Leemos Linea por linea y vamos creando las filas para la respectiva Jtable        
        // Obteniendo la informacion de acuerdo a la posicion en la que se encuentra
        while (scanner.hasNextLine()) {

            String strFila = scanner.nextLine();
            String dniClienteActual = strFila.substring(1, 20);
            dniClientes[j] = dniClienteActual;
            int cantFacturasCliente = 0;

            // CALCULO PARA SACAR BIEN EL MONTO                      
            String montoEnteros = strFila.substring(49, 58);
            String montoDecimales = strFila.substring(58, 60);
            String pesosStr = montoEnteros + "." + montoDecimales;
            double pesos = Double.parseDouble(pesosStr);
            this.total += pesos;
            //FIN DEL CALCULO

            char[] caracteres = strFila.toCharArray();

            String idFactura = strFila.substring(20, 40);

            char[] idFacturaChar = idFactura.toCharArray();

            for (int i = 0; i < dniClientes.length; i++) {
                if (dniClienteActual.equals(dniClientes[i])) {
                    cantFacturasCliente++;
                }
            }
            String strCantFact = valueOf(cantFacturasCliente);
            idFacturaChar[7] = strCantFact.charAt(0);                      

            char[] fechaVto3 = new char[8];
            for (int i = 0; i < fechaVto3.length; i++) {
                fechaVto3[i] = caracteres[i + 60];
            }

            for (int i = 0; i < fechaVto3.length; i++) {
                if (fechaVto3[4] == '1') {
                    switch (fechaVto3[5]) {
                        case '0':
                            fechaVto3[5] = '2';
                            break;
                        case '1':
                            fechaVto3[4] = '0';
                            fechaVto3[3] = '3';
                            break;
                        case '2':
                            fechaVto3[4] = '0';
                            fechaVto3[3] = '3';
                            break;
                    }
                    break;

                } else if (fechaVto3[4] == '0') {
                    switch (fechaVto3[5]) {
                        case '1':
                            fechaVto3[5] = '3';
                            break;
                        case '2':
                            fechaVto3[5] = '4';
                            break;
                        case '3':
                            fechaVto3[5] = '5';
                            break;
                        case '4':
                            fechaVto3[5] = '6';
                            break;
                        case '5':
                            fechaVto3[5] = '7';
                            break;
                        case '6':
                            fechaVto3[5] = '8';
                            break;
                        case '7':
                            fechaVto3[5] = '9';
                            break;
                        case '8':
                            fechaVto3[5] = '0';
                            fechaVto3[4] = '1';
                            break;
                        case '9':
                            fechaVto3[5] = '1';
                            fechaVto3[4] = '1';
                            break;
                    }
                }

            }

            char[] monto = new char[10];
            for (int i = 0; i < monto.length; i++) {
                monto[i] = caracteres[i + 69];
            }

            String msj = strFila.substring(136, 139) + strFila.substring(143, 177) + "0000";
            char[] msjTicket = msj.toCharArray();

            for (int i = 25; i < msjTicket.length; i++) {
                msjTicket[i] = ' ';
            }

            String cod = strFila.substring(188, 249);
            char[] codBarra = cod.toCharArray();

            int r = 0, h = 0, k = 0, o = 0, q = 0;

            for (int i = 162; i < caracteres.length; i++) {
                caracteres[i] = ' ';
            }
            
            for(int i = 0;i < idFacturaChar.length; i++){
                caracteres[i+20] = idFacturaChar[i];
            }

            for (int i = 79; i < caracteres.length; i++) {

                if (i < 87) {
                    caracteres[i] = fechaVto3[r];
                    r++;
                }

                if (i >= 88 && i <= 97) {
                    caracteres[i] = monto[h];
                    h++;
                }

                if (i >= 136 && i < 176) {
                    caracteres[i] = msjTicket[k];
                    k++;
                }

                if (i >= 176 && i <= 182) {
                    caracteres[i] = msjTicket[o];
                    o++;
                }

                if (i > 190 && i <= 250) {
                    caracteres[i] = codBarra[q];
                    q++;
                }

                if (i > 250) {
                    caracteres[i] = '0';
                }

            }

            this.nuevaLinea = valueOf(caracteres);
            this.escritor.escribir(this.nuevaLinea, true);

//                String nuevaLinea2 = valueOf(monto);
//                String nuevaLinea = valueOf(msjTicket);
            System.out.println(this.nuevaLinea);
            area.append(this.nuevaLinea + "\n");
//            FrmHome.LineasAreaTxt.append(this.nuevaLinea + " \n");
            j++;

        }

        scanner.close();
        System.out.println("\n\n FIN DEL PROCESADO");
        System.out.println("EL MONTO CORRECTO ES: " + BigDecimal.valueOf(this.total));
        jText.setText("EL MONTO CORRECTO ES: " + BigDecimal.valueOf(this.total));
    }

}
