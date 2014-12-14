// XML Formatter by MSXML
//
// Reference:
// http://stackoverflow.com/questions/164575/msxml-from-c-pretty-print-indent-newly-created-documents

#import <msxml6.dll>
#include <comip.h>
#include <iostream>

#include "FileStream.hpp"


inline HRESULT raise_on_failure(HRESULT hResult) {
	if (FAILED(hResult))
		_com_raise_error(hResult);

	return hResult;
}


class com {
public:
	com() {
		raise_on_failure(CoInitializeEx(0, COINIT_APARTMENTTHREADED));
	}

	explicit com(DWORD dwCoInit) {
		raise_on_failure(CoInitializeEx(0, dwCoInit));
	}

	~com() {
		CoUninitialize();
	}
};


void WriteXmlWithFormat(MSXML2::IXMLDOMDocument3* document, IStream* output)
{
	MSXML2::IMXWriterPtr writer(__uuidof(MSXML2::MXXMLWriter60));
	writer->omitXMLDeclaration = VARIANT_FALSE;
	writer->standalone = VARIANT_TRUE;
	writer->indent = VARIANT_TRUE;
	writer->byteOrderMark = VARIANT_FALSE;
	writer->encoding = L"utf-8";

	MSXML2::ISAXContentHandlerPtr ch = writer;
	MSXML2::ISAXErrorHandlerPtr eh = writer;
	MSXML2::ISAXDTDHandlerPtr dh = writer;

	MSXML2::ISAXXMLReaderPtr reader(__uuidof(MSXML2::SAXXMLReader60));

	raise_on_failure(reader->putContentHandler(ch));
	raise_on_failure(reader->putErrorHandler(eh));
	raise_on_failure(reader->putDTDHandler(dh));
	wchar_t lexical_handler[] = L"http://xml.org/sax/properties/lexical-handler";
	wchar_t declaration_handler[] = L"http://xml.org/sax/properties/declaration-handler";
	raise_on_failure(reader->putProperty(reinterpret_cast<unsigned short*>(lexical_handler), _variant_t(writer.GetInterfacePtr())));
	raise_on_failure(reader->putProperty(reinterpret_cast<unsigned short*>(declaration_handler), _variant_t(writer.GetInterfacePtr())));

	writer->output = _variant_t(output);
	raise_on_failure(reader->parse(_variant_t(document)));
}



int wmain(int argc, wchar_t const* argv[])
{
	if (argc < 3) {
		std::cerr << "usage: FormatXML.exe {INPUT_FILE} {OUTPUT_FILE}" << std::endl;
		return EXIT_FAILURE;
	}

	com c;

	try {
		bstr_t const infile(argv[1]);
		bstr_t const outfile(argv[2]);

		MSXML2::IXMLDOMDocument3Ptr document(__uuidof(MSXML2::DOMDocument60));
		document->load(infile);

		IStreamPtr output;
		FileStream::OpenFile(outfile, &output, true);

		WriteXmlWithFormat(document, output);
	} catch (_com_error const& e) {
		wchar_t const* const s = e.ErrorMessage();
		std::wcerr << s << std::endl;
	}

	return EXIT_SUCCESS;
}

