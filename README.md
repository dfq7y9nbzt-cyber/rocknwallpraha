# Rock'n'Wall Praha z.s.

Jednoduchy staticky web pro klubove lezecke krouzky Rock'n'Wall Praha z.s.

## Publikace na GitHub Pages

1. Nahrajte soubory do GitHub repozitare.
2. V nastaveni repozitare otevrite `Settings -> Pages`.
3. Zvolte publikaci z vetve `main` a slozky `/root`.
4. Do pole `Custom domain` zadejte `rocknwallpraha.cz`.
5. Po propsani DNS zapnete `Enforce HTTPS`.

## DNS pro FORPSI

Pro hlavni domenu:

```text
A  @  185.199.108.153
A  @  185.199.109.153
A  @  185.199.110.153
A  @  185.199.111.153
```

Pro `www`:

```text
CNAME  www  uzivatel.github.io
```

Hodnotu `uzivatel.github.io` nahradte nazvem GitHub uctu nebo organizace.
