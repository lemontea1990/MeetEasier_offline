module.exports = function ro_list (callback) {

  // modules -------------------------------------------------------------------
  var Connection = require('tedious').Connection;
  var Request = require('tedious').Request;
  var TYPES = require('tedious').TYPES;
  var async = require('async');
  
  // Create connection to database
  var config = {
    userName: 'name', // username of DB
    password: 'pwd', // PWD of DB
    server: 'server addres', // server address
    options: {
      database: 'db name' 
    }
  }

  var connection = new Connection(config);
  connection.on('connect', function(err) {
    if (err) {
      console.log(err);
    } else {
      console.log('Connected');
  
      // Execute all functions in the array serially
      async.waterfall([
        function Start(callback) {
          console.log('Starting...');
          callback(null, 'Jake', 'United States');
        },
        
        function db_Read() {
          console.log('Reading rows from the Table...');
  
          // Read all rows from table
          // new requestion to get the latest information from DB for the data which is currently avaliable and order by time schedule 
          request = new Request(
          'SELECT * FROM Meeting_Status where Subject not like \'%Canceled%\' and Subject not like \'Abg%\' and Subject not like N\'%已取消%\' order by Room_Name, Start_Time;',
          function(err, rowCount, rows) {
            if (err) {
              callback(err);
            } else {
              console.log(rowCount + ' row(s) returned and now is:' + new Date(Date.now()) );
              // callback(null);
            }
          });
  
          // Print the rows read
          var roomLists = [];
          var promise = new Promise(function (resolve, reject) {
			var result = ""; request.on('row', function(columns) {
            	columns.forEach(function(column) {
              		if (column.value === null) {
                		console.log('NULL');
              		} else {
						
						//configure the data result string and assign special identifier 
						result += column.value + "&-&";
                		roomLists.push(result);
              		}
            	});
            
            	resolve(roomLists);
            	// console.log("the length of result is:" + roomLists.length);
            	// console.log(result);
            	result = "";
          	});
          });

          promise.then((roomLists)=>{

            //console.log('promise roomlist', roomLists);
          	global.rl = roomLists;
          	return roomLists, global.rl;
          });
  
          // Execute SQL statement
          connection.execSql(request);

          global.rl = roomLists;
          return roomLists;
          
        }, 
      ],)      
    }
  });

};
