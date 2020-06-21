module.exports = function (callback) {

var Connection = require('tedious').Connection;
var Request = require('tedious').Request;
var TYPES = require('tedious').TYPES;
var async = require('async');

var rl = require("./roomlists.js");
rl();


var config = {
  userName: 'name', // username of DB
  password: 'pwd', // PWD of DB
  server: 'server addres', // server address
  options: {
    database: 'db name' 
  }
}

var deteleObject =   function (obj) {
    var uniques = [];
    var stringify = {};
    for (var i = 0; i < obj.length; i++) {
        var keys = Object.keys(obj[i]);
        keys.sort(function(a, b) {
            return (Number(a) - Number(b));
        });
        var str = '';
        for (var j = 0; j < keys.length; j++) {
            str += JSON.stringify(keys[j]);
            str += JSON.stringify(obj[i][keys[j]]);
        }
        if (!stringify.hasOwnProperty(str)) {
            uniques.push(obj[i]);
            stringify[str] = true;
        }
    }
    uniques = uniques;
    return uniques;
}

var asign_appointment = function(room){

    //room.Appointments = [];
    var roomLists = [];
    var roomLists_clean = [];
    roomLists_clean =  global.rl;

    for(var h=0; h < roomLists_clean.length; h++)
    {
        if(h % 10 == 9)
        {
          //console.log(roomLists_clean[h]);
          roomLists.push(roomLists_clean[h])
        }
    }

    //console.log(roomLists);

    //order by start time but we already ordered the data in SQL result, so this function make nothing effect
    // function sort_start_time(a,b){

    //   console.log("a is:" + a.split("&-&")[2] + "and b is: " + a.split("&-&")[3]);
    //   console.log("now is:" + a.split("&-&")[1]);

    //   a = a.split("&-&")[2];
    //   b = b.split("&-&")[2];

    //   if (a < b) {
    //     return 1;
    // } else if (a > b) {
    //     return -1;
    // } else {
    //     return 0;
    // }
    // }

    //roomLists.sort(sort_start_time);

    //console.log("current room for assign apointment is:" + room.Name);
    
    for(var i=0; i < roomLists.length; i++)
    {
      s = roomLists[i].split("&-&");
      if(s[1] == room.RoomAlias){
        //console.log("start to get the detail of :" + room.RoomAlias);
        
          room.Appointments.Subject = s[7];
          room.Appointments.Organizer = s[4];
          room.Appointments.Start = Date.parse(new Date(s[2]));
          room.Appointments.End = Date.parse(new Date(s[3]));
          room.Appointments.Location = s[1];

          now = Date.now();
          st =  Date.parse(new Date(room.Appointments.Start));
          en = Date.parse(new Date(room.Appointments.End));

  

          if(room.Busy )
          { }
          else
          {
            if( st< now && now < en)
              {
                  room.Busy = true;
              }
          }  
      
          room.Appointments.push({
              "Subject": room.Appointments.Subject,
              "Organizer" : room.Appointments.Organizer,
              "Start" : room.Appointments.Start,
              "End"   : room.Appointments.End
          })        
      }
    }
    return room.Appointments;
  }

  // promise: get all room lists
  var getListOfRooms = function () {
    var promise = new Promise(function (resolve, reject) {
      
      var roomLists = [];
      roomLists =  global.rl;
      var roomAddresses = [];     
      
      for(var i =0; i < roomLists.length; i++)
      {
        
        s = roomLists[i].split("&-&");

        let room = {};
        room.Appointments = [];

        //why here use %10, is because we totally have 10 values in a full string.
        switch(i%10){

          case 1: 
            break;
          case 2:
            var start = s[2],
            end = roomLists[3], now = Date.now();
            
            break;
          case 5:
            
          //just make sure the all these kind of variables all defined as roomname
            room.Name = s[1];
            room.email = s[1];
            room.alias = s[1];
            room.RoomAlias = s[1];
            room.Roomlist = s[1];
            roomAddresses.push(room);

            //console.log(room.Name);
            break;

          default:
            //console.log("start doing");
          }          
      }

     roomAddresses = deteleObject(roomAddresses);
      
     for (var k = 0; k < roomAddresses.length ; k++)
      {
         asign_appointment(roomAddresses[k]);
      }
      resolve(roomAddresses);
    })
    return promise;

  };

  // perform promise chain to get rooms
  getListOfRooms()
  // .then(getAppointmentsForRooms)

  .then(function(rooms){
      callback(null, rooms);
  });
};
