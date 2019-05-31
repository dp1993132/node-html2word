const puppeteer = require('puppeteer');
const Koa = require("koa")
const fs = require("fs")
const path = require("path")
const officegen = require("officegen")

const app = new Koa()

const getDocx = (filename) => new Promise((resolve, reject) => {
    fs.readFile(path.join(__dirname, filename), (err, res) => {
        if (err != null) {
            reject(err)
        } else {
            const docname = filename.replace("png", "docx")
            const docfile = fs.createWriteStream(docname)
            const docx = officegen('docx')
            const p = docx.createP()
            p.addImage(path.join(__dirname, filename))
            docx.generate(docfile)
            docfile.on("finish", () => {
                resolve(fs.createReadStream(docname))
            })
        }
    })
})

app.use(async ctx => {
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.goto("http://" + ctx.query.url);
    // await page.screenshot({ path: 'example.png' });
    const name = Math.random() + "_.png"
    await page.screenshot({ path: name })
    await browser.close();
    try {
        ctx.response.set("Content-Disposition", `attachment;filename=${name.replace("png", "docx")}`)

        ctx.body = await getDocx(name)
    } catch (err) {
        ctx.body = err
    }
    fs.unlink(path.join(__dirname, name), err => {
        if (err != null) {
            console.log(err);
        }
    })
    fs.unlink(path.join(__dirname, name.replace("png", "docx")), err => {
        if (err != null) {
            console.log(err);
        }
    })
})

app.listen(3002)