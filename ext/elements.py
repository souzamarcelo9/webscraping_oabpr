"""
Modulo destinado para armazenar os seletores utilizados na Automação. 

XPATH | CSS
"""

XPATH = {
    'dado': '//table/tbody/tr/td[2]/a',
    'ultima_pagina': '//a[span[contains(text(), "ltima")]]',
    'iframe': '//iframe[contains(@title, "challenge")]'
}

CSS = {
    'reCAPTCHA': '[title=reCAPTCHA]',
    'solved': '[aria-checked]',
    'body': '.rc-anchor-content',
    'audio': '#recaptcha-audio-button',
    'audio-source': 'a[class*="audiochallenge"]',
    'input': '#audio-response',
    'verify_btn': '#recaptcha-verify-button',
    'enviar_btn': '[value="ENVIAR"]',
    'imprimir_btn': '[value="Imprimir"]',
    'rows': 'table.table-striped tbody tr'
}