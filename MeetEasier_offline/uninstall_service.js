let Service = require('node-windows').Service;

let svc = new Service({
        name: 'MeetEasier',    //服务名称
        description: 'The Outlook meeting room status web page.', //描述
        script: 'D:\\Meeting_Online\\MeetEasier-master_win_service\\server.js' //nodejs项目要启动的文件路径
    });

svc.on('uninstall',function(){
      console.log('Uninstall complete.');
      console.log('The service exists: ',svc.exists);
    });

svc.uninstall();
  
 