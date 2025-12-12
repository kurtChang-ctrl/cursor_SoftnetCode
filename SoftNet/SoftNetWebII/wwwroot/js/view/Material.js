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
                { data: 'PartNO', orderable: true },
                { data: 'PartName', orderable: true },
                { data: 'Specification', orderable: true },
                { data: 'Model', orderable: true },
                { data: 'Class', orderable: true },
                { data: 'SafeQTY' },
                { data: 'Unit' },
                { data: 'StoreSTime' },
                { data: 'IS_Store_Test' },
                { data: 'APS_Default_MFNO' },
                { data: 'APS_Default_StoreNO' },
                { data: 'PartType', orderable: true },

            ],
            columnDefs: [
                {
                    targets: [0],
                    render: function (data, type, full, meta) {
                        var tmp = full.ServerId + ';' + full.PartNO;
                        return _crud.dtCrudFun(tmp, full.Name, true, true, false);
                    }
                },
                {
                    targets: [5], render: function (data, type, full, meta) {
                        if (full.Class == '1') { return '原物料'; }
                        else if (full.Class == '2') { return '採購件'; }
                        else if (full.Class == '3') { return '委外加工件'; }
                        else if (full.Class == '4') { return '製造半成品'; }
                        else if (full.Class == '5') { return '製造成品'; }
                        else if (full.Class == '6') { return '刀具'; }
                        else if (full.Class == '7') { return '工具製具'; }
                        else { return '未定義'; }
                    }
                },
                {
                    targets: [9], render: function (data, type, full, meta) {
                        if (full.IS_Store_Test == '1') { return '是'; }
                        else { return '否'; }
                    }
                },
                {
                    targets: [12], render: function (data, type, full, meta) {
                        if (full.PartType == '1') { return '虛料'; }
                        else { return '實料'; }
                    }
                },

            ],
        };

        //initial
        _crud.DbKeyNameS = ['ServerId', 'PartNO'];
        _crud.DbKeyName = 'PartNO';
        _crud.init(config);
    },

};