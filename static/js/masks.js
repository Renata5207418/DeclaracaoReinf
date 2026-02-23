document.addEventListener('DOMContentLoaded', () => {
    // Máscara de Moeda (R$)
    const moneyInputs = document.querySelectorAll('.money-mask');
    moneyInputs.forEach(input => {
        input.addEventListener('input', (e) => {
            let value = e.target.value.replace(/\D/g, '');
            value = (Number(value) / 100).toLocaleString('pt-BR', {
                style: 'currency',
                currency: 'BRL'
            });
            e.target.value = value;
        });
    });

    // Máscara de CPF
    const cpfInput = document.getElementById('cpf');
    if(cpfInput) {
        cpfInput.addEventListener('input', (e) => {
            let v = e.target.value.replace(/\D/g, "");
            v = v.replace(/(\d{3})(\d)/, "$1.$2");
            v = v.replace(/(\d{3})(\d)/, "$1.$2");
            v = v.replace(/(\d{3})(\d{1,2})$/, "$1-$2");
            e.target.value = v.substring(0, 14);
        });
    }

    // Máscara de CNPJ Genérica (para qualquer input com a classe ou id especifico)
    const cnpjInputs = document.querySelectorAll('.cnpj-mask, #company_cnpj');
    cnpjInputs.forEach(input => {
        input.addEventListener('input', (e) => {
            let v = e.target.value.replace(/\D/g, "");
            v = v.replace(/^(\d{2})(\d)/, "$1.$2");
            v = v.replace(/^(\d{2})\.(\d{3})(\d)/, "$1.$2.$3");
            v = v.replace(/\.(\d{3})(\d)/, ".$1/$2");
            v = v.replace(/(\d{4})(\d)/, "$1-$2");
            e.target.value = v.substring(0, 18);
        });
    });
});