Docx-to-HTML
============

This bundle is made from two modified open-source classes. The first perpose was to transfer MS Word content to the webpage in comfort.

Converts .docx files to HTML

This is a PHP class that will convert your .docx files to HTML. It is by far not perfect, but will handle most things decently. Even images. Even in different issues of Word.

The main class requires the following:

- [The ZipArchive class](http://php.net/manual/en/class.ziparchive.php)
- [SimpleXML](http://php.net/manual/en/book.simplexml.php)


### How to use

The sample you can find inside index.php file.

The images' folder should have the same structure as "$dir_container/media" and all links && images would be saved inside media dir. That's it just because at the moment I'm too lazy to upgrade this thing)

### Licence

You can do what you want. I've made this repo just because I hadn't found any moretheless complete solution in php like this one.

Would be nice you make pull requests in case you upgrade this bundle for your perposes.