#include <BLEDevice.h>
#include <BLEUtils.h>
#include <BLEScan.h>
#include <BLEAdvertisedDevice.h>

int scanTime = 1; // In seconds
BLEScan* pBLEScan;
bool isScanning = false; // Default to NOT scanning - CHANGED THIS
bool singleScanMode = false;

// Variables to track the device with highest RSSI (closest device)
BLEAdvertisedDevice* highestRssiDevice = nullptr;
int highestRssi = -200; // Start with a very low value

class MyAdvertisedDeviceCallbacks: public BLEAdvertisedDeviceCallbacks {
    void onResult(BLEAdvertisedDevice advertisedDevice) {
      int currentRssi = advertisedDevice.getRSSI();
      
      // If this is the first device or has higher RSSI than current highest
      if (highestRssiDevice == nullptr || currentRssi > highestRssi) {
        // Update highest RSSI values
        highestRssi = currentRssi;
        
        // Create a new device object to store the details
        if (highestRssiDevice != nullptr) {
          delete highestRssiDevice;
        }
        highestRssiDevice = new BLEAdvertisedDevice(advertisedDevice);
      }
    }
};

void setup() {
  Serial.begin(115200);
  Serial.println("ESP32-C3 BLE Scanner - Ready for Commands");
  Serial.println("Commands: START, STOP, SINGLE");
  
  // Initialize variables
  highestRssiDevice = nullptr;
  highestRssi = -200;

  BLEDevice::init("");
  pBLEScan = BLEDevice::getScan();
  pBLEScan->setAdvertisedDeviceCallbacks(new MyAdvertisedDeviceCallbacks());
  pBLEScan->setActiveScan(true);
  pBLEScan->setInterval(100);
  pBLEScan->setWindow(99);
  
  // Start in stopped state
  isScanning = false;
  Serial.println("SCANNING_STOPPED"); // Send initial state
}

void performScan() {
  // Reset tracking for each new scan
  if (highestRssiDevice != nullptr) {
    delete highestRssiDevice;
    highestRssiDevice = nullptr;
  }
  highestRssi = -200;
  
  BLEScanResults* foundDevices = pBLEScan->start(scanTime, false);
  
  // Send formatted data for Python decoding
  if (highestRssiDevice != nullptr) {
    String deviceName = highestRssiDevice->haveName() ? 
                       highestRssiDevice->getName().c_str() : "Unknown";
    String macAddress = highestRssiDevice->getAddress().toString().c_str();
    int rssi = highestRssiDevice->getRSSI();
    
    // Format: START,device_name,mac_address,rssi,END
    Serial.print("START,");
    Serial.print(deviceName);
    Serial.print(",");
    Serial.print(macAddress);
    Serial.print(",");
    Serial.print(rssi);
    Serial.println(",END");
  } else {
    // No devices found
    Serial.println("START,NoDevice,00:00:00:00:00:00,0,END");
  }
  
  pBLEScan->clearResults();
}

void loop() {
  // Check for incoming commands from Python
  if (Serial.available()) {
    String command = Serial.readStringUntil('\n');
    command.trim();
    
    if (command == "START") {
      isScanning = true;
      singleScanMode = false;
      Serial.println("SCANNING_STARTED");
    } 
    else if (command == "STOP") {
      isScanning = false;
      singleScanMode = false;
      Serial.println("SCANNING_STOPPED");
    }
    else if (command == "SINGLE") {
      isScanning = true;
      singleScanMode = true;
      Serial.println("SINGLE_SCAN");
    }
  }
  
  if (isScanning) {
    performScan();
    
    // If in single scan mode, stop after one scan
    if (singleScanMode) {
      isScanning = false;
      singleScanMode = false;
      Serial.println("SCANNING_STOPPED");
    } else {
      // Continuous mode - wait before next scan
      delay(500);
    }
  } else {
    // Not scanning - small delay to prevent busy waiting
    delay(100);
  }
}