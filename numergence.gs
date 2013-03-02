/*
 * Numergence script
 * Please see http://numergence.com/guide.html for instructions
 *
 * (c) Pavel Suchmann 2013
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 * THE SOFTWARE.
 */

function onOpen() {
  var entries = [
    { name : "Numerge formula", functionName : "menuTask" },
    { name : "User token", functionName : "menuToken" },
    { name : "Settings", functionName : "menuSettings" }
  ];
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.addMenu("Numergence", entries); 
};

// token

function menuToken() {
  var app = UiApp.createApplication()
                 .setTitle('Enter your user token:')
                 .setHeight(130);
  
  var token = app.createTextBox().setWidth(400)
                 .setId('token').setName('token')
                 .setValue(UserProperties.getProperty('numergenceToken'));
  
  var panelV = app.createVerticalPanel();
  panelV.add(token);

  var btnpanel = app.createHorizontalPanel()
                    .setStyleAttribute('marginTop', '10px')
                    .setStyleAttribute('marginLeft', '-5px')
                    .setStyleAttribute('marginBottom', '30px');
  btnpanel.add(app.createButton('Save', app.createServerHandler('onSaveToken').addCallbackElement(token)));
  btnpanel.add(app.createButton('Cancel', app.createServerHandler('onClose')));
  panelV.add(btnpanel);
  
  panelV.add(app.createAnchor("Need help? Follow the guide.", "http://numergence.com/guide.html"));  
  app.add(panelV);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.show(app);  
}

function onSaveToken(eventInfo) {
  var app = UiApp.getActiveApplication();
  UserProperties.setProperty('numergenceToken', eventInfo.parameter.token);
  app.close();
  return app;
}

function clearToken() {
  UserProperties.deleteProperty('numergenceToken'); 
}

// settings

function menuSettings() {
  var app = UiApp.createApplication()
                 .setTitle('Functions used')
                 .setHeight(130)
                 .setWidth(250);
  
  var panel = app.createVerticalPanel();
  panel.add(app.createCheckBox('exp').setValue(true).setEnabled(false));
  panel.add(app.createCheckBox('cos').setValue(true).setEnabled(false));
  panel.add(app.createCheckBox('sin').setValue(true).setEnabled(false));
  panel.add(app.createCheckBox('ln').setValue(true).setEnabled(false));             
  app.add(panel);
  
  var btnpanel = app.createHorizontalPanel()
                    .setHorizontalAlignment(UiApp.HorizontalAlignment.RIGHT)
                    .setWidth(250);
  btnpanel.add(app.createButton('Close', app.createServerHandler('onClose')));
  app.add(btnpanel);
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.show(app);   
}

function onClose() {
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

// task

function menuTask() {
  var token = UserProperties.getProperty('numergenceToken');
  if (token) {
    task();
  } else {
    menuToken();
  }
}

function task() { 
  // task status
  var state = loadState();
  var activeTask = isActiveTask(state['task']);
  //Logger.log("activeTask = %s state = %s", Utilities.jsonStringify(activeTask), Utilities.jsonStringify(state));
  if (activeTask) {
    Browser.msgBox('Numergence', 'You have an active task. Please wait to sync with the server.', Browser.Buttons.OK);
  }
    
  var app = UiApp.createApplication().setTitle('Numergence').setHeight(480);

  var panel1 = app.createVerticalPanel();
  var btnpanel = app.createHorizontalPanel();
  
  // status  
  var status = app.createLabel('Status: not yet started').setStyleAttribute('fontSize', 'larger').setId('status');  
  panel1.add(status);
  
  var rng = app.createLabel('Selection: ' + SpreadsheetApp.getActiveSheet().getActiveSelection().getA1Notation()).setId('rng');
  panel1.add(rng);

  // timer
  var handler = app.createServerHandler('handleTimer');
  var chk = app.createCheckBox('tick').addValueChangeHandler(handler).setVisible(false).setId('chk').setValue(state.tick == 1);
  panel1.add(chk);
  
  // buttons
  var selbutton = app.createButton('Start task', app.createServerHandler('onStart'));
  selbutton.setId('selbutton').setEnabled(!activeTask);

  var stopbutton = app.createButton('Stop task', app.createServerHandler('onStop'));
  stopbutton.setId('stopbutton').setEnabled(activeTask);

  var newbutton = app.createButton('New task', app.createServerHandler('onNew'));
  newbutton.setId('newbutton').setEnabled(false);
  
  var handlerSelBtn = app.createClientHandler()
    .forEventSource().setEnabled(false)
    .forTargets(stopbutton).setEnabled(true)
    .forTargets(status).setText('Status: starting');
  selbutton.addClickHandler(handlerSelBtn);

  var handlerStopBtn = app.createClientHandler()
    .forEventSource().setEnabled(false)
    .forTargets(status).setText('Status: stopping');
  stopbutton.addClickHandler(handlerStopBtn);
  
  btnpanel.add(selbutton);
  btnpanel.add(stopbutton);
  btnpanel.add(newbutton);
  btnpanel.add(app.createButton('Close view', app.createServerHandler('onClose')));
  panel1.add(btnpanel);
  app.add(panel1);
  
  // result 
  var panelRes = app.createVerticalPanel().setId('results').setVisible(false);
  panelRes.add(app.createLabel('Results').setStyleAttribute('fontSize', 'larger').setStyleAttribute('marginTop', '10px'));
    
  var scroll = app.createScrollPanel().setId('scroll');
  scroll.add(app.createGrid().setId('grid').setBorderWidth(1).setCellPadding(3).setCellSpacing(0));
  scroll.setPixelSize(500, 360);
  panelRes.add(scroll);
  
  var panel3 = app.createHorizontalPanel();
  panel3.add(app.createLabel('').setId('steps').setStyleAttribute('marginTop', '5px'));
  panel3.add(app.createLabel('').setId('evaluations').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '20px'));
  panelRes.add(panel3);
  app.add(panelRes); 

  // continuation
  if (activeTask) {
    handleTimer();
  }  
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.show(app);  
}

function onStart() {
  var app = UiApp.getActiveApplication();
  
  var rng = SpreadsheetApp.getActiveSheet().getActiveSelection();
  app.getElementById('rng').setText('Selection: ' + rng.getA1Notation());
  
  var values = rng.getValues();
  var valErr = validateRange(values);
  if (valErr) {
    app.getElementById('status').setText('Status: incorrect input: ' + valErr);
    app.getElementById('stopbutton').setEnabled(false);
    app.getElementById('newbutton').setEnabled(true);
    return app;
  }
  
  var state = loadState();
  state['out_range'] = {
    row: rng.getRow(),
    column: rng.getColumn() + rng.getWidth(),
    height: rng.getHeight()
  };
  saveState(state);
  
  sendRequest(app, 'post', '/tasks', { data: Utilities.jsonStringify(values) }); 
  return app;
}

function onStop() {
  var app = UiApp.getActiveApplication();
  var state = loadState();
  sendRequest(app, 'put', '/tasks/' + state.task.id, 'command=stop&outcome="stopped by user"');
  return app;
}

function onNew() {
  clearState();
  var app = UiApp.getActiveApplication();
  app.close();
  menuTask();
  return app;
}

function handleTimer() {
  var app = UiApp.getActiveApplication(); 
  var state = loadState();
  if (state['task']) {
    sendRequest(app, 'get', '/tasks/' + state.task.id + '?r=' + Math.round(Math.random()*100000), '');
  }
  return app;
}

function validateRange(values) {
  if (values.length > 20) {
    return 'Max. number of rows exceeded';
  }
  if (values.length < 2) {
    return 'At least 2 rows necessary';
  }
  
  var row = values[0];
  if (row.length > 5) {
    return 'Max. number of columns exceeded';
  }
  if (row.length < 2) {
    return 'At least one input and one output columns necessary';
  }

  for (var i=0; i<values.length; i++) {
    var row = values[i];
    for (var j=0; j<row.length; j++) {
      if (typeof row[j] !== 'number') {
        return 'Non-numeric cell(s) in selected range';
      }
    }
  }
  return null;
}

function sendRequest(app, method, path, payload) {
  var options = {
    headers: { 'X-Numergence-Token' : UserProperties.getProperty('numergenceToken') },
    method: method,
    payload: payload
  };
  var result = UrlFetchApp.fetch('http://api.numergence.com/client' + path, options);
  var output = Utilities.jsonParse(result.getContentText());
  if (output['success']) {       
    render(app, output['task']);
  } else {
    Browser.msgBox('ERROR: ' + output['message']);
  }
}

function render(app, task) {
  var state = loadState();
  state['task'] = task;
  state.tick = 1 - state.tick;
  if (isActiveTask(task)) {
    saveState(state);
  }
  
  if (task.result && task.result['results'] && state['out_range']) {
    app.getElementById('steps').setText('Step: ' + task.result.step);
    app.getElementById('evaluations').setText(' Evaluations: ' + task.result.numof_evaluations);

    var results = task.result.results;
    var grid = app.getElementById('grid');
    grid.resize(0, 0).resize(results.length+1, 4)
      .setText(0, 0, 'use').setStyleAttributes(0, 0, {"backgroundColor": "#E0E0E0", "textAlign": "center", "fontSize": "small"})
      .setText(0, 1, 'RMSE').setStyleAttributes(0, 1, {"backgroundColor": "#E0E0E0", "textAlign": "center", "fontSize": "small"})
      .setText(0, 2, 'cmplx').setStyleAttributes(0, 2, {"backgroundColor": "#E0E0E0", "textAlign": "center", "fontSize": "small"})
      .setText(0, 3, 'formula').setStyleAttributes(0, 3, {"backgroundColor": "#E0E0E0", "textAlign": "left", "fontSize": "small"});
    var radioHandler = app.createServerHandler('radioHandler');
    for (var i=0; i<results.length; i++) {
      var radio = app.createRadioButton('radio').setId('radio_'+i).setValue(i == 0).addClickHandler(radioHandler);
      var rmsd = Math.round(Math.sqrt(results[i].error / state.out_range.height)*1000)/1000;
      grid.setWidget(i+1, 0, radio)
        .setText(i+1, 1, rmsd).setStyleAttribute(i+1, 1, "fontSize", "small")
        .setText(i+1, 2, results[i].complexity).setStyleAttribute(i+1, 2, "fontSize", "small")
        .setText(i+1, 3, results[i].formula).setStyleAttribute(i+1, 3, "fontSize", "small");
    }
    applyFormula(0); 
    
    //app.getElementById('scroll').setPixelSize(500, 340);
    app.getElementById('results').setVisible(true);
  } else {
    app.getElementById('results').setVisible(false);
  }

  var statusText = 'Status: ' + task.status;
  if (task['outcome']) {
    statusText += ' Outcome: ' + task.outcome;
  }
  app.getElementById('status').setText(statusText);
  
  if (isActiveTask(task)) {
    // timer
    Utilities.sleep(12000);  
    app.getElementById('chk').setValue((state.tick == 1),true);
  } else {
    app.getElementById('newbutton').setEnabled(true);  
    app.getElementById('stopbutton').setEnabled(false);
  }
  return app;
}

function loadState() {
  var stateText = UserProperties.getProperty('numergenceState');
  var state = {};
  if (stateText) {
    state = Utilities.jsonParse(stateText);
  } 
  if (!(state['tick'])) {
    state['tick'] = 0;
  }
  return state;
}

function saveState(state) {
  UserProperties.setProperty('numergenceState', Utilities.jsonStringify(state)); 
}

function clearState() {
  UserProperties.deleteProperty('numergenceState'); 
}

function isActiveTask(task) {
  if (task && task['status'] && task.status != 'finished' && task.status != 'stopped') { 
    return true; 
  } else { 
    return false;
  }
}

function radioHandler(eventInfo) {
  applyFormula(eventInfo.parameter.source.replace('radio_',''));
}

function applyFormula(index) {
  var state = loadState();
  if (state['out_range'] && state['task'] && state.task['result'] && state.task.result['results']) {
    var rng = SpreadsheetApp.getActiveSheet().getRange(state.out_range.row, state.out_range.column, state.out_range.height, 1);
    var results = state.task.result.results;
    if (results.length > index) { 
      rng.setFormulaR1C1(results[index].formula);
    }
  }
}

