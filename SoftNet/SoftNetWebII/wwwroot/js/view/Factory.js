var _me = {
    hasStaionStatus: false,
    hasSimulatioStatus: false,
    STView2Work_PageReload: false,
    hasStaionChangQTY: false,

    init: function () {
        //datatable config
        var config = {
            columns: [
                { data: '_Crud' },
                { data: 'FactoryName', orderable: true },
                { data: 'Manager', orderable: true },
                { data: 'Address', orderable: true },
                { data: 'Telephone', orderable: true },
            ],
            columnDefs: [
                {
                    targets: [0], render: function (data, type, full, meta) {
                        return _crud.dtCrudFun(full.Id, full.Name, true, true, false);
                    }
                },
            ],
        };

        //initial
        _crud.DbKeyName = 'Id';
        _crud.init(config);
    },

};