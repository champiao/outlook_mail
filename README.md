# OutlookMail

OutllokMail é um script python que tem como objetivo logar em uma conta outlook e baixar PDFs com os emails da imbox como conteúdo.

## Utilization

create a .env File to configure your credentials
```bash
EMAIL=yourOutlookEmail@example.com
PASS=YOURpassword
```

## Usage

```python
pip install -r requirement.txt
python3 app.py
```

## Resultado

Este código te gera duas pastas separados/ onde cada email estará em um arquivo pdf com o nome email_{identificador}.pdf e a pasta unificados/ onde ele te trará um arquivo pdf com todos os emails e com o nome Final-{timeStamp}.pdf

## License

[CHAMPIAO](https://champiao.com.br)
