function app(prefix) {
    var self = this;
    self.inputs = ko.observableArray();
    self.outputs = ko.observableArray();
    self.prefix = prefix;
    self.Calculate = function () {
        var inputdata = {};
        for (var i in self.inputs()) {
            inputdata[self.inputs()[i].name] = { value: self.inputs()[i].value() };
        }
        $.ajax({
            url: '/api/sheets/' + prefix + '/output',
            method: 'post',
            contentType: "application/json; charset=utf-8",
            dataType: "json",
            data: JSON.stringify(inputdata),
            success: function (result) {
                for (var i in self.outputs())
                    self.outputs()[i].value(result[self.outputs()[i].name].value);
            }
        });

    }
    self.UpdateInput = function () {
        var inputdata = {};
        for (var i in self.inputs()) {
            inputdata[self.inputs()[i].name] = { value: self.inputs()[i].value() };
        }
        $.ajax({
            url: '/api/sheets/' + prefix + '/input',
            method: 'post',
            data: JSON.stringify(inputdata),
            dataType: "json",
            contentType: "application/json; charset=utf-8",
            success: function (result) {
                for (var i in self.inputs())
                    self.inputs()[i].valid(result[self.inputs()[i].name].valid);
            }
        });

    }
    $.ajax({
        url: '/api/sheets/' + prefix + '/input',
        success: function (result) {
            for (var input in result) {
                var inputobject = {
                    name: input,
                    value: ko.observable(result[input].value),
                    valid: ko.observable(result[input].valid),
                    options: result[input].options,
                    enabled: ko.observable(result[input].enabled),
                    errormessage: result[input].errormessage,
                    unit: result[input].unit
                };
                inputobject.value.subscribe(function () {
                    self.UpdateInput();
                });
                self.inputs.push(inputobject);
            }
        }
    });
    $.ajax({
        url: '/api/sheets/' + prefix + '/output',
        success: function (result) {
            for (var output in result)
                self.outputs.push({
                    name: output,
                    value: ko.observable(result[output].value),
                    unit: result[output].unit,
                    description: result[output].description
                });
        }
    });
}