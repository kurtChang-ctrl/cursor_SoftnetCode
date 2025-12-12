var _me = {
    hasStaionStatus: false,
    hasSimulatioStatus: false,
    STView2Work_PageReload: false,
    hasStaionChangQTY: false,

    init: function () {        
        //datatable config
        var config = {
            columns: [
                { data: 'Code' },
                { data: 'Name' },
                //{ data: 'Icon' },
                { data: 'Url' },
                { data: 'Sort' },
                { data: '_Fun' },
            ],
            columnDefs: [
				{ targets: [4], render: function (data, type, full, meta) {
                    return _crud.dtCrudFun(full.Id, full.Name, true, true, true);
                }},
            ],
        };

        //initial
        _me.mRoleProg = new EditMany('Id', 'eformRoleProg', 'tplRoleProg', 'tr');
		_crud.init(config, [null, _me.mRoleProg]);
    },

}; //class