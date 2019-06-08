let backUrl = '/sysConfig/';

let admin = new Vue({
    el: '#admin',
    data: {
        // 表格表头
        columns: [
            {
                type: 'selection',
                width: 60,
                align: 'center'
            },
            {
                title: '程序名',
                key: 'name'
            },
            {
                title: '概述',
                key: 'description'
            },
            {
                title: '命令',
                key: 'cmd'
            }
        ],
        // 表格内容
        list: [],
        // 模态框状态
        modal:false,
        modalEdit:false,
        name:'',
        cmd:'',
        description:'描述内容……',
        totalData:null,
        curId:'',
        selection:[],
    },
    // 页面初始化
    mounted: function () {
        axios.get(backUrl + 'getListAdmin',{
            params: {
                page:0,
                count:6,
            }
        }).then(res => {
            let that = this;
            let data = res.data.data;
            let page_info = res.data.page_info;

            that.list = data;
            that.totalData = page_info.total;
            that.nextStatus = page_info.has_next;
        }).catch(res => {

        })
    },
    methods: {
        // 双击弹出编辑
        handleEdit (index) {
            this.modalEdit = true;
            this.name = index.name;
            this.description = index.description;
            this.cmd = index.cmd;
            this.curId = index.id;
        },
        // 选择删除项
        handleDelete (selection) {
            let that = this;
            that.selection = selection;
        },
        // 确认新增识别
        onAdd () {
            let that = this;
            axios.post(backUrl + 'add',{
                name:that.name,
                description:that.description,
                cmd:that.cmd
            }).then(res => {
                location.reload();
            })
        },
        // 确认编辑
        onEdit () {
            let that = this;
            axios.post(backUrl + 'updateById',{
                name:that.name,
                description:that.description,
                cmd:that.cmd,
                id:that.curId
            }).then(res => {
                location.reload()
            })
        },
        // 确认删除
        onDelete () {
            let that = this;
            let idList = [];
            for(let i = 0;i<that.selection.length;i++){
                let id = that.selection[i].id;
                idList.push(id);
            }
            axios.post(backUrl + 'deleteByIds', 
                idList
            ).then(res => {
                location.reload();
            })
        },
        // 分页
        onChangePage (val) {
            let that = this;
            axios.get(backUrl + 'getListAdmin',{
                params: {
                    page:val - 1,
                    count:6,
                }
            }).then(res => {
                let data = res.data.data;
                that.list = data;
            }).catch(res => {

            })
        }
    },
})