import 'dart:io';

import 'package:external_path/external_path.dart';
import 'package:syncfusion_flutter_xlsio/xlsio.dart' as xcel;
import 'package:flutter/material.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Flutter Demo',
      theme: ThemeData(
        // This is the theme of your application.
        //
        // TRY THIS: Try running your application with "flutter run". You'll see
        // the application has a blue toolbar. Then, without quitting the app,
        // try changing the seedColor in the colorScheme below to Colors.green
        // and then invoke "hot reload" (save your changes or press the "hot
        // reload" button in a Flutter-supported IDE, or press "r" if you used
        // the command line to start the app).
        //
        // Notice that the counter didn't reset back to zero; the application
        // state is not lost during the reload. To reset the state, use hot
        // restart instead.
        //
        // This works for code too, not just values: Most code changes can be
        // tested with just a hot reload.
        appBarTheme: AppBarTheme(
            backgroundColor: Theme.of(context).colorScheme.background,
            titleTextStyle: TextStyle(
                fontSize: 26,
                color: Theme.of(context).colorScheme.onPrimaryContainer)),
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.deepPurple),
        useMaterial3: true,
      ),
      home: const HomeScreen(),
    );
  }
}

class HomeScreen extends StatelessWidget {
  const HomeScreen({Key? key}) : super(key: key);

  @override
  Widget build(BuildContext context) {
    final xcel.Workbook workbook = xcel.Workbook();
    final xcel.Worksheet sheet = workbook.worksheets[0];

    List myList = [
      {
        "title": "How to Add .env File in Flutter?",
        "link": "https://www.geeksforgeeks.org/how-to-add-env-file-in-flutter/"
      },
      {
        "title": "Flutter - Select Single and Multiple Files From Device",
        "link":
            "https://www.geeksforgeeks.org/flutter-select-single-and-multiple-files-from-device/"
      },
      {
        "title": "Autofill Hints Suggestion List in Flutter",
        "link":
            "https://www.geeksforgeeks.org/autofill-hints-suggestion-list-in-flutter/"
      },
      {
        "title": "How to Integrate Razorpay Payment Gateway in Flutter?",
        "link":
            "https://www.geeksforgeeks.org/how-to-integrate-razorpay-payment-gateway-in-flutter/"
      },
      {
        "title": "How to Setup Multiple Flutter Versions on Mac?",
        "link":
            "https://www.geeksforgeeks.org/how-to-setup-multiple-flutter-versions-on-mac/"
      },
      {
        "title": "How to Change Package Name in Flutter?",
        "link":
            "https://www.geeksforgeeks.org/how-to-change-package-name-in-flutter/"
      },
      {
        "title":
            "Flutter - How to Change App and Launcher Title in Different Platforms",
        "link":
            "https://www.geeksforgeeks.org/flutter-how-to-change-app-and-launcher-title-in-different-platforms/"
      },
      {
        "title": "Custom Label Text in TextFormField in Flutter",
        "link":
            "https://www.geeksforgeeks.org/custom-label-text-in-textformfield-in-flutter/"
      }
    ];
    return Scaffold(
      appBar: AppBar(
        title: const Text('Flutter export to excel'),
      ),
      body: SingleChildScrollView(
        child: Column(
          children: List.generate(
              myList.length,
              (index) => Card(
                    color: Theme.of(context).colorScheme.secondaryContainer,
                    child: ListTile(
                      title: Text(
                        '${myList[index]['title']}',
                        style: const TextStyle(color: Colors.pink),
                      ),
                      subtitle: Text('${myList[index]['link']}'),
                    ),
                  )),
        ),
      ),
      floatingActionButtonLocation:
          FloatingActionButtonLocation.miniCenterFloat,
      floatingActionButton: FloatingActionButton.extended(
        onPressed: () async {
          sheet.getRangeByIndex(1, 1).setText("Title");
          sheet.getRangeByIndex(1, 2).setText("Link");

          for (var i = 0; i < myList.length; i++) {
            final item = myList[i];
            sheet.getRangeByIndex(i + 2, 1).setText(item["title"].toString());
            sheet.getRangeByIndex(i + 2, 2).setText(item["link"].toString());
          }
          final List<int> bytes = workbook.saveAsStream();
          var directory = await ExternalPath.getExternalStoragePublicDirectory(
              ExternalPath.DIRECTORY_DOCUMENTS);
          File("$directory/excel.xlsx")
            ..createSync(recursive: true)
            ..writeAsBytes(bytes ?? []);

        },
        label: const Text("Create Excel sheet"),
      ),
    );
  }
}
