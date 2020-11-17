const XLSX = require('./lib/xlsx.full.min.js');
const fs = require('fs')

var allExsServices = []

const main = async () => {

    const { Sheets } = XLSX.readFile('./Organigrammes MCN.xlsx');

    let allNodes = []
    let allTrees = {}
    let allOrgKeys = []

    for (const sheetKey of Object.keys(Sheets)) {
        const jsonKey = sheetKey.includes('-') ? sheetKey.split('-')[0].replace('Organigramme ', '') : sheetKey.replace('Organigramme ', '')
        allOrgKeys.push(jsonKey)
        const Sheet = Sheets[sheetKey];
        const lines = XLSX.utils.sheet_to_csv(Sheet, { FS: "|", blankrows: false }).split('\n').splice(0)

        const nodes = lines.reduce((acc, curr, index) => {
            const node = lineToJsonData(curr, jsonKey, index)
            if (node) {
                acc = { ...acc, [index]: node }
            }
            return acc
        }, {})

        const [childrens, nodesWithPathAndId] = buildChildrens(nodes)

        allExsServices = allExsServices.map(el => {
            if (!el.new.id) {
                el.new = nodesWithPathAndId[el.new].obj
            }
            return el
        })

        allNodes = [...allNodes, ...Object.keys(nodesWithPathAndId).map(el => nodesWithPathAndId[el])]

        const tree = []
        const roots = Object.keys(nodesWithPathAndId).filter(key => nodesWithPathAndId[key].level == 0)

        for (const root of roots) {
            const subTree = buildSubTree(root, JSON.parse(JSON.stringify(nodesWithPathAndId)), childrens)
            tree.push(subTree)
        }

        allTrees = Object.assign(allTrees, { [jsonKey]: { name: "My Company Name", shortName: "MCN", path: "MCN", children: tree } })

    }

    fs.writeFileSync('./charts.json', JSON.stringify(allTrees, null, 4))

    let dicName = {}
    for (const n of allNodes) {
        if (dicName[n.obj.name]) {
            dicName[n.obj.name].push(n.obj.id)
        } else {
            dicName = { ...dicName, [n.obj.name]: [n.obj.id] }
        }
    }

    const autocompletev = Object.keys(dicName).map(el => {
        const entry = dicName[el]
        return {
            name: el,
            shortName: entry[0].split('/').slice(-1)[0],
            orgKey: entry.map(ele => { return ele.split(':')[0] })
        }
    }).sort((a, b) => {
        if (a.name < b.name) {
            return -1;
        }
        if (a.name > b.name) {
            return 1;
        }
        return 0;
    })

    const getName = (id, allNodes) => {
        for (const node of allNodes) {
            if (id == node.obj.id) {
                return node.obj.name
            }
        }
    }

    const changelog = allExsServices.map(el => {
        el.old = el.old.map(o => {
            const orgKeyN = el.new.orgKey
            const exOrgKey = allOrgKeys[allOrgKeys.indexOf(orgKeyN) - 1]
            o.orgKey = exOrgKey
            o.id = exOrgKey + ":" + o.path
            o.name = getName(o.id, allNodes)
            return o
        })
        return el
    })

    for (const autoc of autocompletev) {
        let changes = []
        for (const change of changelog) {
            if (change.new.name == autoc.name) {
                changes.push(change)
            }
        }
        if (changes.length > 0) {
            autoc.changes = changes
        }
    }

    fs.writeFileSync('./autocompletev.json', JSON.stringify(autocompletev, null, 4))

    fs.writeFileSync('./changelog.json', JSON.stringify(changelog, null, 4))

    fs.writeFileSync('./nodes.json', JSON.stringify(allNodes, null, 4))

    fs.writeFileSync('./nameIds.json', JSON.stringify(dicName, null, 4))

}

const buildSubTree = (nodeKey, nodes, childrens) => {
    const node = nodes[nodeKey].obj
    const children = childrens[nodeKey]
    if (!children) { return node }
    node.children = children.map(child => buildSubTree(child, nodes, childrens))
    return node
}

const buildChildrens = (nodes) => {
    let nodesWithPathAndId = {}
    let childrens = {}
    let path = []
    for (const nodeKey of Object.keys(nodes)) {
        const node = nodes[nodeKey]
        if (path.length == 0 || node.level == 0) {
            path = [nodeKey]
        } else {
            let parentNodeKey
            if (node.level < path.length - 1) {
                //Le noeud est fils du node.level-1 element
                parentNodeKey = path[node.level - 1]
                path[node.level] = nodeKey
                path = path.slice(0, node.level + 1)
            } else if (node.level > path.length - 1) {
                //Le noeud est fils du dernier element
                parentNodeKey = path[path.length - 1]
                path.push(nodeKey)
            } else {
                //Le noeud est fils de l'avant dernier element
                parentNodeKey = path[path.length - 2]
                path[path.length - 1] = nodeKey
            }
            updateChildrens(childrens, parentNodeKey, nodeKey)
        }
        const nodePath = "MCN/" + path.map(el => nodes[el].obj.shortName).join('/')
        const nodeId = node.obj.orgKey + ":" + nodePath
        nodesWithPathAndId = {
            ...nodesWithPathAndId, [nodeKey]: {
                level: node.level,
                obj: {
                    orgKey: node.obj.orgKey,
                    name: node.obj.name,
                    shortName: node.obj.shortName,
                    comment: node.obj.comment,
                    path: nodePath,
                    id: nodeId
                }
            }
        }
    }
    return [childrens, nodesWithPathAndId]
}

const updateChildrens = (childrens, parentNodeKey, childKey) => {
    if (Object.keys(childrens).includes(parentNodeKey)) {
        childrens[parentNodeKey].push(childKey)
    } else {
        childrens = Object.assign(childrens, { [parentNodeKey]: [childKey] })
    }
}

const lineToJsonData = (line, orgKey, i) => {
    const spl = line.split('|')
    let level = 0
    let name = ""
    let shortName = ""
    let exServices = []
    let comment = ""
    for (const s of spl) {
        if (s == "") {
            level++
        } else {

            let reg
            
			[shortName, name] = s.includes(":") ? s.split("(")[0].split(":").map(el => el.replace('"', '').trim()) : ["", s]

			let fusionNService = new RegExp('.+\\\(ex ([A-Z1-9-/ +]+)\\\)', 'g')
			let exService = new RegExp('.+\\\(ex ([A-Z1-9-/]+)\\\)', 'g')

			exServices = []
			reg = exService.exec(s)
			if (reg) {
				exServices.push(reg[1])
			} else {
				reg = fusionNService.exec(s)
				if (reg) {
					const spl = reg[1].split(' + ')
					for (const sp of spl) {
						exServices.push(sp)
					}
				}
			}
            if (spl[13]) {
                comment = spl[13]
            }
            break
        }

    }

    const obj = { name: name, shortName: shortName != "" ? shortName : undefined, orgKey: orgKey, comment: comment != "" ? comment : undefined }
    if (obj.name != "") {
        if (exServices.length > 0) {
            allExsServices.push({
                date: orgKey,
                type: exServices.length == 1 ? "rename" : "fusion",
                old: exServices.map(el => { return { path: "MCN/" + el, shortName: el.split('/').slice(-1)[0] } }),
                new: i
            })
        }
        return { level: level, obj: obj }
    }

    return null

}

main()