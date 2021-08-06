package main

/*
==========================================================================

 	PACKAGE NAME: TaniumGet

	AUTHOR: Jim Borecky , Homelab
	DATE  : 1/30/2020

	COMMENT: This is just an example of functions that can access the Tanium
			infrastucture

==========================================================================
*/

import (
	"bytes"
	"crypto/tls"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
)

//GetSessionID is used to extract a SessionKey
//from a JSON
func GetSessionID(body string) string {
	//////// Attempting to parse JSON here
	//    {
	//		"data":{
	//	     		"session":"2-59005-b5f80502123155cf68d2aab4613b9c2c4ef5c118019e63d8f9c0a6212c2cfc1b"
	//	  		   }
	//	  }

	type Session struct {
		SessionKey string `json:"session"`
	}
	type SessionAnswer struct {
		Data *Session `json:"data"`
	}
	var myReturnSession SessionAnswer

	//fmt.Println(string(body))
	json.Unmarshal([]byte(body), &myReturnSession)
	//_ = json.NewDecoder(body).Decode(&myReturnSession)
	return string(myReturnSession.Data.SessionKey)
	//fmt.Println(string(myReturnSession.Data.SessionKey))
}

//LoginToTanium is a function to log on to
//Tanium returning a session key
func LoginToTanium(username string, domain string, password string) string {
	//Create Login credentials
	type Login struct {
		UserName string `json:"username"`
		Domain   string `json:"domain"`
		Password string `json:"password"`
	}

	MyJSON := &Login{username, domain, password} //Load credentials

	buf := new(bytes.Buffer)            //Setup buffer for transport
	json.NewEncoder(buf).Encode(MyJSON) //Encode buffer using structure

	//Setup request
	req, err := http.NewRequest("POST", "https://tanium1.rougeone.borecky.net/api/v2/session/login", buf)
	req.Header.Set("Content-Type", "application/json")

	//Disable cert check
	tr := &http.Transport{TLSClientConfig: &tls.Config{InsecureSkipVerify: true}}
	client := &http.Client{Transport: tr}

	//Go baby GO!
	resp, err := client.Do(req)

	//Check for errors
	if err != nil {
		log.Fatalln(err)
	}

	//Close connection
	defer resp.Body.Close()

	//Grab the body that was returned and extract pending no errors
	body, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		log.Fatalln(err)
	}

	//log.Println(string(body))

	return GetSessionID(string(body))
}

func main() {

	SessionKey := LoginToTanium("myID", "rougeone", "Mypassword")
	fmt.Println("Session ID =", SessionKey)
	//Function New question
}
