let app = new Vue({
    el: '#app',
    data: {
        // 选择active状态
        activeClass: 0,
        // 顶部导航栏列表
        navItems: [],
        title: '通用文字',
        itemDescription: '识别文档的所有信息',
        cmd: '',
        uploadUrl: '',
        imgSrc:'',
        uploadStatus:-1,
    },

    methods: {
        handleError () {
            this.$Notice.warning({
                title: '温馨提醒',
                desc: '请上传png，jpeg，jpg和pdf格式文件'
            });
        },
        beforeUpload(file) { /*上传前确定上传地址*/
            let that = this;
            uploadPic(that.uploadUrl, this);
        },

        // index：navItems索引,遍历
        selectSort(index) {
            let that = this;
            console.log(index)
            let itemId = this.navItems[index].id;
            this.uploadStatus = itemId;
            this.activeClass = index;
            that.uploadStatus = -1;
            that.imgSrc = null;
            axios.get('/sysConfig/getOneById', {
                params: {
                    id: itemId
                }
            }).then(res => {
                let title = res.data.data.name;
                let itemDescription = res.data.data.description;
                let cmd = res.data.data.cmd;
                let targetJson = document.getElementById('targetjson');
                targetJson.value = '';
                that.title = title;
                that.itemDescription = itemDescription;
                that.cmd = cmd;
                that.uploadUrl = '/python/exePython?cmd=' + cmd

            })
        }
    },
    mounted: function () {
        //导航
        let that = this;
        axios.get('/sysConfig/getListIndex').then(res => {
            console.log(res)
            let data = res.data.data;
            for (let i = 0; i < data.length; i++) {
                that.navItems.push(data[i])
            }
            that.uploadUrl = '/python/exePython?cmd=' + data[0].cmd;
        }).catch(error => {
            console.log(error);
        })
    }

})

function uploadPic(url, _this) {
    var form = document.getElementsByClassName('ivu-upload-input');
    var formData = new FormData();
    formData.append("file", form[0].files[0]);
    var ii = layer.load();
    axios.post(url, formData).then(function (response) {
        app.uploadStatus = 1;
        let imgUrl = response.data.msg;
        app.imgSrc = imgUrl;

        if (response.data.code !== 2000) {
            document.getElementById("sourcejson").innerText = JSON.stringify(response.data);
            formatJson()
            layer.close(ii);
            _this.$Message.error('服务器异常！请重试或联系管理员！');
        } else {
            console.log(JSON.stringify(response.data.data))
            document.getElementById("sourcejson").innerText = JSON.stringify(response.data);
            formatJson()
            layer.close(ii);
        }
    }).catch(function (error) {
        document.getElementById("sourcejson").innerText = error;
        formatJson();
        layer.close(ii);
        _this.$Message.error(error);
        console.log('======识别失败==========【' + error + '】')
    });
}

function repeat(s, count) {
    return new Array(count + 1).join(s);
}

function formatJson() {

    var json = document.form1.sourcejson.value;

    var i = 0,
        len = 0,
        tab = "    ",
        targetJson = "",
        indentLevel = 0,
        inString = false,
        currentChar = null;


    for (i = 0, len = json.length; i < len; i += 1) {
        currentChar = json.charAt(i);

        switch (currentChar) {
            case '{':
            case '[':
                if (!inString) {
                    targetJson += currentChar + "\n" + repeat(tab, indentLevel + 1);
                    indentLevel += 1;
                } else {
                    targetJson += currentChar;
                }
                break;
            case '}':
            case ']':
                if (!inString) {
                    indentLevel -= 1;
                    targetJson += "\n" + repeat(tab, indentLevel) + currentChar;
                } else {
                    targetJson += currentChar;
                }
                break;
            case ',':
                if (!inString) {
                    targetJson += ",\n" + repeat(tab, indentLevel);
                } else {
                    targetJson += currentChar;
                }
                break;
            case ':':
                if (!inString) {
                    targetJson += ": ";
                } else {
                    targetJson += currentChar;
                }
                break;
            case ' ':
            case "\n":
            case "\t":
                if (inString) {
                    targetJson += currentChar;
                }
                break;
            case '"':
                if (i > 0 && json.charAt(i - 1) !== '\\') {
                    inString = !inString;
                }
                targetJson += currentChar;
                break;
            default:
                targetJson += currentChar;
                break;
        }
    }
    document.form1.targetjson.value = targetJson;
    return;
}