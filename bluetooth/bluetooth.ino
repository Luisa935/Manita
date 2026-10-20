HardwareSerial BluetoothSerial(2);

void setup() {
  Serial.begin(115200);(*baudrate mismo del monitor serial*)

  BluetoothSerial.begin(9600, SERIAL_8N1, 16, 17);

  Serial.println("ESP32 lista");
  Serial.println("Escribe algo para enviarlo al HC-05");
}

void loop() {

  // Computadora -> HC-05
  while (Serial.available()) {
    char c = Serial.read();
    BluetoothSerial.write(c);
  }

  // HC-05 -> Computadora
  while (BluetoothSerial.available()) {
    char c = BluetoothSerial.read();
    Serial.write(c);
  }
}