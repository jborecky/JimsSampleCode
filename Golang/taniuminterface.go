package main

/*
==========================================================================

 	PACKAGE NAME: TaniumInterface

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
	"net/url"
	"strconv"
	"time"
)

//GetSessionID is used to extract a SessionKey
//from a JSON
func GetSessionID(body string) string {
	type Session struct {
		SessionKey string `json:"session"`
	}
	type SessionAnswer struct {
		Data *Session `json:"data"`
	}
	var myReturnSession SessionAnswer

	json.Unmarshal([]byte(body), &myReturnSession)
	return string(myReturnSession.Data.SessionKey)
} //End of GetSessionID

//LoginToTanium :
//
//is a function to log on to Tanium returning a session key
func LoginToTanium(username string, domain string, password string) string {
	//Create Login credentials
	type Login struct {
		UserName string `json:"username"`
		Domain   string `json:"domain"`
		Password string `json:"password"`
	}

	//Build MyJSON
	MyJSON := &Login{username, domain, password} //Load credentials

	//From here just formatted info to pass to the Connector
	if b, err := json.Marshal(MyJSON); err == nil {
		body := Connect(b, "session/login", "")
		return GetSessionID(string(body))
	}
	return "error"
}

//CreateQuestion -- Get's question by query
func CreateQuestion(Question string, SessionKey string) int {
	byt := []byte(`{"question_text" : "Get Computer Name from all machines",
	    "selects" : [
          {
            "group" : {
              "and_flag" : false,
              "deleted_flag" : false,
              "filters" : [],
              "not_flag" : false,
              "sub_groups" : []
            },
            "sensor" : {
              "hash" : 3409330187,
              "name" : "Computer Name"
            }
          }
        ],
        "sensor_references" : [
          {
            "name" : "Computer Name",
            "start_char" : "4"
          }
        ]
     }`)

	body := Connect(byt, "questions", SessionKey)

	type Session struct {
		SessionID int `json:"id"`
	}
	type SessionAnswer struct {
		Data *Session `json:"data"`
	}
	var myReturnSession SessionAnswer

	json.Unmarshal([]byte(body), &myReturnSession)
	return myReturnSession.Data.SessionID
} //End of Create Question

//GetSensorByName --:  Blah
//	TODO blah blah blah
func GetSensorByName(SensorName string, SessionKey string) string {
	//NO STRUCTURE TEST
	byt := []byte(``)

	body := Connect(byt, "sensors/by-name/"+url.PathEscape(SensorName), SessionKey)
	return string(body)
} //End of GetSensorByName

//GetQuestionByID --: blah blah blah
func GetQuestionByID(QuestionID int, SessionKey string) string {
	//https://localhost/api/v2/questions/3
	byt := []byte(``)
	body := Connect(byt, "questions/"+strconv.Itoa(QuestionID)+"?json_pretty_print=1", SessionKey)
	return string(body)
}

//GetQuestionResultsByID --: blah blah blah
func GetQuestionResultsByID(QuestionID int, SessionKey string) string {
	//https://localhost/api/v2/result_data/question/33?
	byt := []byte(``)
	body := Connect(byt, "result_data/question/"+strconv.Itoa(QuestionID)+"?json_pretty_print=1", SessionKey)
	return string(body)
}

// Connect function
func Connect(b []byte, URLTail string, SessionKey string) string {
	//Setup request
	URL := "https://tanium1.rougeone.borecky.net/api/v2/" + URLTail

	//Check for json
	if len(b) != 0 {
		var NewJSON map[string]interface{}
		if err := json.Unmarshal(b, &NewJSON); err != nil {
			fmt.Println("Life sucks!")
		}
		buf := new(bytes.Buffer)             //Setup buffer for transport
		json.NewEncoder(buf).Encode(NewJSON) //Encode buffer using structure
		req, err := http.NewRequest("POST", URL, buf)
		req.Header.Set("Content-Type", "application/json")
		if SessionKey != "" {
			req.Header.Set("session", SessionKey)
		}

		//Disable cert check
		tr := &http.Transport{TLSClientConfig: &tls.Config{InsecureSkipVerify: true}}
		client := &http.Client{Transport: tr}

		//Go baby GO!
		resp, err := client.Do(req)

		//Check for errors
		if err != nil {
			log.Fatalln(err)
		}

		//Closes reader
		defer resp.Body.Close()

		//Grab the body that was returned and extract pending no errors
		body, err := ioutil.ReadAll(resp.Body)
		if err != nil {
			log.Fatalln(err)
		}
		return string(body)
	}

	req, err := http.NewRequest("GET", URL, nil)

	if SessionKey != "" {
		req.Header.Set("session", SessionKey)
	}

	//Disable cert check
	tr := &http.Transport{TLSClientConfig: &tls.Config{InsecureSkipVerify: true}}
	client := &http.Client{Transport: tr}

	//Go baby GO!
	resp, err := client.Do(req)

	//Check for errors
	if err != nil {
		log.Fatalln(err)
	}

	//Closes reader
	defer resp.Body.Close()

	//Grab the body that was returned and extract pending no errors
	body, err := ioutil.ReadAll(resp.Body)
	if err != nil {
		log.Fatalln(err)
	}
	return string(body)

}

func main() {

	SessionKey := LoginToTanium("myID", "rougeone", "myPassword")
	fmt.Println("Session ID =", SessionKey)
	fmt.Println("========================================================")
	fmt.Println(GetSensorByName("Computer Name", SessionKey))
	fmt.Println("========================================================")
	questionID := CreateQuestion("MyQuestion", SessionKey)
	fmt.Println(questionID)
	fmt.Println("Waiting 30 seconds to gather machines")
	time.Sleep(30 * time.Second)
	fmt.Println(GetQuestionByID(questionID, SessionKey))
	fmt.Println(GetQuestionResultsByID(questionID, SessionKey))

}
