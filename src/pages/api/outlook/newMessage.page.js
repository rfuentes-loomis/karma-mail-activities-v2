export default async function handler(req, res) {
  if (req.method === "POST") {
    // Handle POST request
    // Parse the JSON body
    const body = JSON.parse(req.body);

    // Access the parsed body
    console.log("Parsed body:", body);
    //sample value
    // {"value":[{"changeType":"created","clientState":null,"resource":"Users/653dbcfa-205d-4ce2-9909-7706ef55c3a4/Messages/AAMkAGYwZGQ3NzBkLTE4MTMtNGE5NC04M2Q2LWM0YWU3MmU0NjY3ZgBGAAAAAACATjC0X7t_T5XcS4DSj4NGBwDl3Ngp3uykQ6Y_LAOtyIq2AAAAAAEKAADl3Ngp3uykQ6Y_LAOtyIq2AAA709TcAAA=","resourceData":{"@odata.etag":"W/\"CQAAABYAAADl3Ngp3uykQ6Y+LAOtyIq2AAA7tYs+\"","@odata.id":"Users/653dbcfa-205d-4ce2-9909-7706ef55c3a4/Messages/AAMkAGYwZGQ3NzBkLTE4MTMtNGE5NC04M2Q2LWM0YWU3MmU0NjY3ZgBGAAAAAACATjC0X7t_T5XcS4DSj4NGBwDl3Ngp3uykQ6Y_LAOtyIq2AAAAAAEKAADl3Ngp3uykQ6Y_LAOtyIq2AAA709TcAAA=","@odata.type":"#Microsoft.Graph.Message","id":"AAMkAGYwZGQ3NzBkLTE4MTMtNGE5NC04M2Q2LWM0YWU3MmU0NjY3ZgBGAAAAAACATjC0X7t_T5XcS4DSj4NGBwDl3Ngp3uykQ6Y_LAOtyIq2AAAAAAEKAADl3Ngp3uykQ6Y_LAOtyIq2AAA709TcAAA="},"subscriptionExpirationDateTime":"2024-02-22T16:08:53.238+00:00","subscriptionId":"61b190fe-042c-4e15-86a2-b1d270eba2b4","tenantId":"6703241b-3ef6-4e56-b8ca-a04eb0dd9ac7"}]}

    //value[0].resourceData.id
    res.status(200).json();
  } else {
    // Handle other HTTP methods
    res.status(405).json({ message: "Method Not Allowed" });
  }
}
