import re

def refactor_laws():
    with open('templates/laws.html', 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Replace everything from start up to <!-- HERO -->
    # We want to keep the inner <style> but modify it.
    style_match = re.search(r'<style>(.*?)</style>', content, re.DOTALL)
    style_content = style_match.group(1) if style_match else ""

    # Map colors in style_content
    style_content = style_content.replace('var(--dark)', 'var(--bg)')
    style_content = style_content.replace('var(--card)', 'var(--surface)')
    style_content = style_content.replace('var(--red)', 'var(--primary)')
    style_content = style_content.replace('var(--border)', 'rgba(255,255,255,0.08)')
    style_content = style_content.replace('.sidebar{', '.laws-sidebar{')
    style_content = style_content.replace('.sidebar{display:none}', '.laws-sidebar{display:none}')
    style_content = style_content.replace('body{', '.laws-body-wrapper{') # Prevent body overrides

    # Build the new header
    header = """{% extends "base.html" %}
{% block title %}Corporate Laws Hub — MyBBT{% endblock %}
{% block page_id %}laws-hub{% endblock %}
{% block topbar_title %}Laws Hub{% endblock %}
{% block topbar_breadcrumb %}<span>Tools</span> › Corporate Laws{% endblock %}

{% block extra_css %}
<style>
""" + style_content + """
</style>
{% endblock %}

{% block content %}
<div class="laws-body-wrapper">
"""

    # Extract the main body part starting from <!-- HERO -->
    hero_start = content.find('<!-- HERO -->')
    if hero_start == -1:
        print("Could not find HERO")
        return
        
    main_body = content[hero_start:]
    
    # Extract javascript
    script_match = re.search(r'<script>(.*?)</script>', main_body, re.DOTALL)
    script_content = script_match.group(1) if script_match else ""
    
    # Remove javascript from main_body, remove closing body/html
    main_body = re.sub(r'<script>.*?</script>', '', main_body, flags=re.DOTALL)
    main_body = main_body.replace('</body>', '').replace('</html>', '')
    
    # Replace sidebar class
    main_body = main_body.replace('class="sidebar"', 'class="laws-sidebar"')

    # Build footer
    footer = """
</div>
{% endblock %}

{% block extra_js %}
<script>
""" + script_content + """
</script>
{% endblock %}
"""

    new_content = header + main_body + footer

    with open('templates/laws.html', 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    print("Successfully refactored laws.html")

if __name__ == "__main__":
    refactor_laws()
