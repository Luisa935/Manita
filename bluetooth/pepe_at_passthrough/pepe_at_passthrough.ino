/*
  ============================================================================
  Puente USB <-> PEPE para mandarle comandos AT (cambiar su baud rate)
  ============================================================================

  Para que sirve:
    Al revisar el firmware original de la manita (lehand.ino), PEPE esta
    cableado a los pines 11 (RX) y 12 (TX) de la placa -- son pines de
    "Bluetooth por software" (SoftwareSerial), NO los pines 0/1 que usa el
    puerto USB. Por eso, para mandarle comandos AT a PEPE (por ejemplo para
    subirle el baud rate a 57600) hace falta este sketch puente: recibe por
    USB lo que escribes en el Monitor Serie y se lo reenvia a PEPE por los
    pines 11/12, y viceversa.

  Pasos:
    1. Sube este sketch a la placa de la manita (con el IDE de Arduino,
       seleccionando la placa correcta -- la misma que usaste para ver que
       el chip es un MEGA328P).
    2. Desconecta la alimentacion. Manten presionado el botoncito de PEPE
       (o pon su pin KEY/EN a 3.3V si no tiene boton) y vuelve a conectar
       la alimentacion -- asi entra en modo de comandos AT. El LED de PEPE
       deberia parpadear lento (cada ~2 segundos) en vez de rapido.
    3. Abre el Monitor Serie del IDE de Arduino a 38400 baudios (es el
       baud fijo del modo AT de los HC-05, sin importar a que este
       configurado el modo de datos), con el final de linea en
       "Ambos NL y CR".
    4. Escribe "AT" y dale enter -- deberia responder "OK". Si no responde
       nada, revisa que SI este en modo AT (paso 2) y que el Monitor Serie
       este a 38400 baudios.
    5. Escribe "AT+UART?" para ver el baud actual.
    6. Escribe "AT+UART=57600,0,0" para dejarlo en 57600 (o el baud que
       hayas elegido en lehand_30hz.ino, deben coincidir). Responde "OK".
    7. Desconecta y reconecta la alimentacion SIN mantener presionado el
       boton esta vez, para volver al modo de datos normal.
    8. Sube lehand_30hz.ino (con Bth.begin(57600) puesto igual) a la placa.
  ============================================================================
*/

#include <SoftwareSerial.h>

#define BTH_RX 11
#define BTH_TX 12

SoftwareSerial Bth(BTH_RX, BTH_TX);

void setup() {
  Serial.begin(38400);  // baud fijo del modo AT de PEPE
  Bth.begin(38400);
}

void loop() {
  if (Bth.available()) {
    Serial.write(Bth.read());
  }
  if (Serial.available()) {
    Bth.write(Serial.read());
  }
}
