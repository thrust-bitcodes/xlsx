var Files = Java.type("java.nio.file.Files");
var Paths = Java.type("java.nio.file.Paths");

let majesty = require('majesty')
let xlsx = require('../dist/index.js');

function exec(describe, it, beforeEach, afterEach, expect, should, assert) {

    describe("Testes de modelagem de xlsx", function () {
        var dtNasc = new Date(2018, 4, 18);

        var rows = [{
            id: 1,
            nome: 'Zé',
            dt_nasc: dtNasc,
            salario: 1122.5
        }, {
            id: 2,
            nome: 'Jão',
            dt_nasc: dtNasc,
            salario: 965.7
        }, {
            id: 3,
            nome: 'Maria',
            dt_nasc: dtNasc,
            salario: 1100.0
        }];

        it("Escrita/Leitura simples de planilha", function () {
            var bytes = xlsx.create(rows);
            var result = xlsx.read(bytes);

            expect(result.length, 'Deve conter a mesma quantidade de linhas')
                .to.equals(rows.length);

            expect(result[0].id).to.equals(1);
            expect(result[0].nome).to.equals('Zé');
            expect(result[0].dt_nasc.getTime()).to.equals(dtNasc.getTime());
            expect(result[0].salario).to.equals(1122.5);
        });

        it("Escrita/Leitura com metadados", function () {
            var metadata = {
                hasHeader: true,
                headerStyle: {
                    bold: true,
                    horizontalAlignment: 'center',
                    fontName: 'Colibri'
                },
                style: {
                    fontName: 'Courrier New',
                },
                columns: {
                    id: {
                        description: 'Código',
                        italic: true,
                    },
                    nome: {
                        doubleUnderline: true
                    },
                    dt_nasc: {
                        description: 'Dt Nasc.',
                        striked: true,
                    },
                    salario: {
                        type: 'currency'
                    }
                }
            }

            var bytes = xlsx.create(rows, metadata);
            var result = xlsx.read(bytes);

            expect(result.length, 'Deve conter a mesma quantidade de linhas')
                .to.equals(rows.length);

            expect(result[0]['Código']).to.equals(1);
            expect(result[0].nome).to.equals('Zé');
            expect(result[0]['Dt Nasc.'].getTime()).to.equals(dtNasc.getTime());
            expect(result[0].salario).to.equals(1122.5);

            expect(1).to.equals(1);
        })
    });
}

function readAllBytes(path) {
    return Files.readAllBytes(Paths.get(path));
}

let res = majesty.run(exec)

print(res.success.length, " scenarios executed with success and")
print(res.failure.length, " scenarios executed with failure.\n")

res.failure.forEach(function (fail) {
    print("[" + fail.scenario + "] =>", fail.execption)
})

exit(res.failure.length);