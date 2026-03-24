require 'docx'
require 'pp'

doc = Docx::Document.open('spec/fixtures/with-images.docx')

# 先看有哪些图片关系
pp doc.images
# 例如输出: { "rId5" => "word/media/image1.png" }

# 方式A：按 rId 替换
doc.replace_image('rId5', 'spec/fixtures/replacement.png')

# 方式B：按 entry 路径替换（两种二选一即可）
# doc.replace_image('word/media/image1.png', 'spec/fixtures/replacement.png')

doc.save('tmp/with-images-edited.docx')
puts 'done'