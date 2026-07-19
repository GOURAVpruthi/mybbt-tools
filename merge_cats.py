import re

with open('templates/pdf_tools.html', 'r', encoding='utf-8') as f:
    html = f.read()

# Pattern to find and remove:
#       </div>
#     </div>
#
#     <!-- CATEGORY NAME -->
#     <div class="category-section">
#       <div class="category-title">Category Name</div>
#       <div class="tools-grid">
pattern = re.compile(r'\s*</div>\s*</div>\s*<!--.*?-->\s*<div class="category-section">\s*<div class="category-title">.*?</div>\s*<div class="tools-grid">', re.DOTALL)

new_html = pattern.sub('', html)

# We also need to remove the extra </div></div> at the very end of the tools-container, 
# because we removed the opening tags for the category-section and tools-grid.
# Wait, actually, if we just remove the boundaries, they all merge into the first tools-grid.

with open('templates/pdf_tools.html', 'w', encoding='utf-8') as f:
    f.write(new_html)

print("Categories merged successfully.")
