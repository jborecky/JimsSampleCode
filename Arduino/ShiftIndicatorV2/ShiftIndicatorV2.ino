/*========================================================================================================================
*
* Arduino Source File -- Created with Arduino Software version 1.8.19
*
* SKETCH NAME: ShiftIndicatorV2.ino
*
* AUTHOR: Jim Borecky
* DATE  : 05-13-2022
*
* COMMENT: This sketch was written to get inputs from a 4L60E transmission and output the 4 signal input pattern to a 
*   8x8 matix display the shows (P,R,N,O,D,2,1)
*   Original source was ShiftIndicator.ino which will have some corrections that need to be made.
*
* REQUIREMENTS:
*         Written for Ardunio UNO and tested on Ardunio Nano Every
*   INPUT PINS: from transmission(configurable). See source comments.
*   OUTPUT PINS:  GND to - (pin29)
*                 +5V to + (pin27)
*                 A5  to C (pin24)
*                 A4  to D (pin23)
*         
* NOTE: It doesn't appear there is a pin one on the Adafruit Mini 8x8 LED Matrix w/I2C Backpack - Red [ADA870] but can
*       be inserted into the board either way. (Needs to be comfirmed.)
*========================================================================================================================*/
//Configurable Variables
int SignalPin_A = 5;  //~D5(pin8)
int SignalPin_B = 4;  //D4(pin7)
int SignalPin_C = 3;  //~D3(pin6)
int SignalPin_P = 2;  //D2(pin5)

/*Truth table for 4L60E
  
          SigA(Blk/Wht) SigB(Yel) SigC(Gry) SigP(Wht)      Value
  Park      HIGH          LOW       LOW       HIGH           9
  Reverse   HIGH          HIGH      LOW       LOW            12
  Neutral   LOW           HIGH      LOW       HIGH           5
  Overdrive LOW           HIGH      HIGH      LOW            6
  Drive     HIGH          HIGH      HIGH      HIGH           15
  Second    HIGH          LOW       HIGH      LOW            10
  First     LOW           LOW       HIGH      HIGH           3

  */
static const int Park = 9;
static const int Reverse = 12;
static const int Neutral = 5;
static const int Overdrive = 6;
static const int Drive = 15;
static const int Second = 10;
static const int First = 3;


//Variables used to store values
int SignalA = 0;
int SignalB = 0;
int SignalC = 0;
int SignalP = 0;
int leds = 0;

#include <Wire.h>
#include <Adafruit_GFX.h>
#include "Adafruit_LEDBackpack.h"

Adafruit_8x8matrix matrix = Adafruit_8x8matrix();

static const uint8_t PROGMEM
OverDrive_bmp[] =
  { B00111100,
    B01000010,
    B10111001,
    B10100101,
    B10100101,
    B10111001,
    B01000010,
    B00111100 };

void setup() {
  //Setup the transmission inputs
  pinMode(SignalPin_A, INPUT_PULLUP);
  pinMode(SignalPin_B, INPUT_PULLUP);
  pinMode(SignalPin_C, INPUT_PULLUP);
  pinMode(SignalPin_P, INPUT_PULLUP); 

  //Initialize the Matrix
  Serial.begin(9600);
  Serial.println("8x8 LED Matrix Test");
  matrix.begin(0x70);  // pass in the address
}

void loop() {
  // Read the inputs and multiply
  if (digitalRead(SignalPin_A) == LOW) {
    SignalA = 0;
  } else {
    SignalA = 1;
  };
  Serial.print(SignalA);
  if (digitalRead(SignalPin_B) == LOW) {
    SignalB = 0;
  } else {
    SignalB = 2;
  };
  Serial.print(SignalB);
  if (digitalRead(SignalPin_C) == LOW) {
    SignalC = 0;
  } else {
    SignalC = 4;
  };
  Serial.print(SignalC);
  if (digitalRead(SignalPin_P) == LOW) {
    SignalP = 0;
  } else {
    SignalP = 8;
  };
  Serial.print(SignalP);
  leds = SignalA + SignalB + SignalC + SignalP;

  //Setup output based on the multiplier(P,R,N,O,D,2,1)
  matrix.setRotation(0);
  
  Serial.print (" ");
  Serial.println(leds);
    
  switch (leds) {
      case Park: //Park
          matrix.clear();
          matrix.setCursor(1,0);
          matrix.print("P");
          matrix.writeDisplay();
      break;
      case Reverse: //Reverse
          matrix.clear();
          matrix.setCursor(1,0);
          matrix.print("R");
          matrix.writeDisplay();
      break;
      case Neutral: //Neutral
          matrix.clear();
          matrix.setCursor(1,0);
          matrix.print("N");
          matrix.writeDisplay();
      break;
      case Overdrive: //OverDrive
          matrix.clear();
          matrix.drawBitmap(0, 0, OverDrive_bmp, 8, 8, LED_ON);
          matrix.writeDisplay();
      break;
      case Drive: //Drive
          matrix.clear();
          matrix.setCursor(1,0);
          matrix.print("D");
          matrix.writeDisplay();
      break;
      case Second: //Second
          matrix.clear();
          matrix.setCursor(1,0);
          matrix.print("2");
          matrix.writeDisplay();
      break;
      case First: //First
          matrix.clear();
          matrix.setCursor(1,0);
          matrix.print("1");
          matrix.writeDisplay();
      break;
      default: //Error
          for (int8_t x=7; x>=-30; x--) {
              matrix.clear();
              matrix.setCursor(x,0);
              matrix.print("E"); matrix.println(leds);
              //matrix.print(leds);
              //matrix.print(Error);
              matrix.writeDisplay();  
              delay(100);
          }
      break;
   };
}
